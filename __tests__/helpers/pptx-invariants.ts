import JSZip from 'jszip';
import {
  DOMParser,
  Document as XmlDocument,
  Element as XmlElement,
} from '@xmldom/xmldom';

/**
 * ROADMAP Phase 5, Tier 1 — package invariants for every written archive.
 *
 * `errors` are the corruption classes that break decks in PowerPoint and fail
 * the test that wrote the file:
 *
 *  1. a part that is not well-formed XML;
 *  2. an r:id/r:embed/... attribute with no entry in the part's _rels file;
 *  3. a relationship that an attribute actually references whose target part
 *     is missing (a silently dead hyperlink, image or chart);
 *  4. a part not covered by [Content_Types].xml;
 *  5. a ppt/presentation.xml slide-list entry whose slide part is missing.
 *
 * `knownIssues` are pre-existing library behaviors that PowerPoint tolerates.
 * They are reported for future cleanup work but do not fail tests:
 *
 *  - stale relationship entries left behind when copied content is re-targeted
 *    (media dedup, chart/image swaps) — no attribute references them;
 *  - notesSlide relationships to a notesMaster that is never copied;
 *  - slide parts orphaned by removeExistingSlides/removeSlide without
 *    `cleanup`, and media/chart parts nothing references anymore.
 *
 * Escalating a known-issue class into an error must come with the library fix,
 * mirroring the Tier-2 allowlist approach (fail only on *new* error classes).
 */

const RELATIONSHIP_ATTRIBUTE_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

interface Relationship {
  id: string;
  target: string;
  external: boolean;
}

export interface PptxInvariantReport {
  errors: string[];
  knownIssues: string[];
}

export async function checkPptxInvariants(
  data: Buffer,
): Promise<PptxInvariantReport> {
  const errors: string[] = [];
  const knownIssues: string[] = [];
  const archive = await JSZip.loadAsync(data);

  const partNames = Object.keys(archive.files).filter(
    (name) => !archive.files[name].dir,
  );
  const partSet = new Set(partNames);

  // 1. every XML part parses
  const xmlParts = new Map<string, XmlDocument>();
  for (const name of partNames) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    const content = await archive.files[name].async('text');
    try {
      xmlParts.set(name, parseStrict(content));
    } catch (error) {
      errors.push(`${name}: not well-formed XML (${error})`);
    }
  }

  const relsByPart = new Map<string, Map<string, Relationship>>();
  for (const [name, doc] of xmlParts) {
    if (!name.endsWith('.rels')) continue;
    const rels = new Map<string, Relationship>();
    const elements = doc.getElementsByTagName('Relationship');
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      rels.set(element.getAttribute('Id'), {
        id: element.getAttribute('Id'),
        target: element.getAttribute('Target'),
        external: element.getAttribute('TargetMode') === 'External',
      });
    }
    relsByPart.set(relsOwner(name), rels);
  }

  // 2. r:* attribute references resolve in the owning part's rels
  const referencedIdsByPart = new Map<string, Set<string>>();
  for (const [name, doc] of xmlParts) {
    if (name.endsWith('.rels')) continue;
    const rels = relsByPart.get(name);
    const referenced = new Set<string>();
    referencedIdsByPart.set(name, referenced);
    for (const [attribute, value] of relationshipAttributes(doc)) {
      if (value === '') continue;
      referenced.add(value);
      if (!rels?.has(value)) {
        errors.push(
          `${name}: ${attribute}="${value}" has no entry in ${relsPath(name)}`,
        );
      }
    }
  }

  // 3. relationship targets resolve — an error when an attribute references
  //    the relationship, a known issue when the entry is stale
  const targetedParts = new Set<string>();
  for (const [owner, rels] of relsByPart) {
    const referenced = referencedIdsByPart.get(owner);
    for (const rel of rels.values()) {
      if (rel.external) continue;
      const resolved = resolveTarget(owner, rel.target);
      targetedParts.add(resolved);
      if (partSet.has(resolved)) continue;
      if (referenced?.has(rel.id)) {
        errors.push(
          `${relsPath(owner)}: ${rel.id} → "${rel.target}" is referenced in ` +
            `${owner} but resolves to missing part ${resolved}`,
        );
      } else {
        knownIssues.push(
          `${relsPath(owner)}: stale relationship ${rel.id} → "${rel.target}" ` +
            `(missing part ${resolved}, but nothing references ${rel.id})`,
        );
      }
    }
  }

  // orphaned media/chart/embedding parts — tolerated bloat
  for (const name of partNames) {
    if (!/^ppt\/(media|embeddings|charts)\//.test(name)) continue;
    if (name.includes('/_rels/')) continue;
    if (!targetedParts.has(name)) {
      knownIssues.push(`${name}: orphaned part, no relationship targets it`);
    }
  }

  // 4. [Content_Types].xml covers every part
  const contentTypes = xmlParts.get('[Content_Types].xml');
  if (!contentTypes) {
    errors.push('[Content_Types].xml is missing or unparseable');
  } else {
    const defaults = new Set<string>();
    const defaultElements = contentTypes.getElementsByTagName('Default');
    for (let i = 0; i < defaultElements.length; i++) {
      defaults.add(defaultElements[i].getAttribute('Extension').toLowerCase());
    }
    const overrides = new Set<string>();
    const overrideElements = contentTypes.getElementsByTagName('Override');
    for (let i = 0; i < overrideElements.length; i++) {
      overrides.add(overrideElements[i].getAttribute('PartName'));
    }
    for (const name of partNames) {
      if (name === '[Content_Types].xml') continue;
      const extension = name.includes('.')
        ? name.slice(name.lastIndexOf('.') + 1).toLowerCase()
        : '';
      if (!defaults.has(extension) && !overrides.has(`/${name}`)) {
        errors.push(`${name}: not covered by [Content_Types].xml`);
      }
    }
  }

  // 5. presentation.xml slide list ↔ slide parts
  const presentation = xmlParts.get('ppt/presentation.xml');
  const presentationRels = relsByPart.get('ppt/presentation.xml');
  if (presentation && presentationRels) {
    const listedSlides = new Set<string>();
    const slideIds = presentation.getElementsByTagName('p:sldId');
    for (let i = 0; i < slideIds.length; i++) {
      const rId = slideIds[i].getAttribute('r:id');
      const rel = presentationRels.get(rId);
      if (rel && !rel.external) {
        listedSlides.add(resolveTarget('ppt/presentation.xml', rel.target));
      }
    }
    for (const listed of listedSlides) {
      if (!partSet.has(listed)) {
        errors.push(
          `ppt/presentation.xml: slide list entry ${listed} has no part`,
        );
      }
    }
    for (const name of partNames) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
      if (!listedSlides.has(name)) {
        knownIssues.push(
          `${name}: slide part is not listed in ppt/presentation.xml`,
        );
      }
    }
  }

  return { errors, knownIssues };
}

function parseStrict(content: string): XmlDocument {
  return new DOMParser({
    onError: (level, message) => {
      if (level === 'fatalError' || level === 'error') {
        throw new Error(message);
      }
    },
  }).parseFromString(content, 'application/xml');
}

/** `ppt/slides/_rels/slide1.xml.rels` → `ppt/slides/slide1.xml` */
function relsOwner(relsName: string): string {
  return relsName.replace(/_rels\//, '').replace(/\.rels$/, '');
}

function relsPath(partName: string): string {
  const slash = partName.lastIndexOf('/');
  const dir = slash === -1 ? '' : partName.slice(0, slash + 1);
  const base = partName.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

/** Resolve a relationship Target relative to the part that declares it. */
function resolveTarget(owner: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1);
  }
  const base = owner.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') {
      base.pop();
    } else if (segment !== '.' && segment !== '') {
      base.push(segment);
    }
  }
  return base.join('/');
}

/** All attributes in the r: namespace (r:id, r:embed, r:link, r:dm, …). */
function* relationshipAttributes(
  doc: XmlDocument,
): Generator<[string, string]> {
  const walk = function* (node: XmlElement): Generator<[string, string]> {
    for (let i = 0; i < node.attributes.length; i++) {
      const attribute = node.attributes[i];
      if (attribute.namespaceURI === RELATIONSHIP_ATTRIBUTE_NS) {
        yield [attribute.name, attribute.value];
      }
    }
    let child = node.firstChild;
    while (child) {
      if (child.nodeType === 1) {
        yield* walk(child as XmlElement);
      }
      child = child.nextSibling;
    }
  };
  if (doc.documentElement) {
    yield* walk(doc.documentElement);
  }
}
