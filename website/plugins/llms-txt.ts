import * as fs from 'fs';
import * as path from 'path';
import type { LoadContext, Plugin } from '@docusaurus/types';

/**
 * ROADMAP 6.4 — AI rendering of the docs, generated at build time:
 *
 * - a Markdown twin of every HTML page at the same route + `.md`
 *   (guides and the typedoc-generated API reference alike),
 * - `llms.txt`: a flat index of every guide page with its one-line
 *   frontmatter description,
 * - `llms-full.txt`: the guide corpus concatenated into one file.
 *
 * Everything is derived from the docs plugin's loaded content (routes,
 * titles, sidebar order), so a new page only needs its frontmatter and a
 * sidebar id. A guide page without a frontmatter `description` fails the
 * build — that one-liner is the llms.txt payload.
 */

// Minimal structural types for what we read from the docs plugin content —
// narrower than @docusaurus/plugin-content-docs' internals on purpose.
type DocMeta = {
  id: string;
  title: string;
  permalink: string;
  source: string;
  frontMatter: Record<string, unknown>;
};

type SidebarItem = {
  type: string;
  id?: string;
  items?: SidebarItem[];
  link?: { type: string; id?: string };
};

type DocsContent = {
  loadedVersions: Array<{
    docs: DocMeta[];
    sidebars: Record<string, SidebarItem[]>;
  }>;
};

const stripFrontmatter = (raw: string): string => {
  const match = raw.match(/^---\r?\n[\s\S]*?\r?\n---\r?\n?/);
  return match ? raw.slice(match[0].length) : raw;
};

// Twin bodies always start with an H1: pages carry their title in
// frontmatter, and only some repeat it as a Markdown heading.
const normalizeBody = (raw: string, title: string): string => {
  let body = stripFrontmatter(raw).trim();
  if (!body.startsWith('# ')) {
    body = `# ${title}\n\n${body}`;
  }
  // category-index pages have their twin at <dir>.md, not <dir>/index.md
  // (mirrors the route: /api → api.html / api.md)
  return body.replace(/\]\((\.\/[^)#]+)\/index\.md(#[^)]*)?\)/g, ']($1.md$2)');
};

// Sidebar ids in display order (categories contribute their link doc first).
const flattenSidebarIds = (items: SidebarItem[]): string[] =>
  items.flatMap((item) => {
    if ((item.type === 'doc' || item.type === 'ref') && item.id) {
      return [item.id];
    }
    if (item.type === 'category') {
      const linkId =
        item.link?.type === 'doc' && item.link.id ? [item.link.id] : [];
      return [...linkId, ...flattenSidebarIds(item.items ?? [])];
    }
    return [];
  });

export default function llmsTxtPlugin(context: LoadContext): Plugin<undefined> {
  return {
    name: 'llms-txt',

    async postBuild({ outDir, siteConfig, baseUrl, plugins }) {
      const docsPlugin = plugins.find(
        (p) => p.name === 'docusaurus-plugin-content-docs',
      );
      if (!docsPlugin) {
        throw new Error('llms-txt: docusaurus-plugin-content-docs not found');
      }
      const version = (docsPlugin.content as DocsContent).loadedVersions[0];

      // site root as absolute URL, with trailing slash
      const siteRoot = siteConfig.url + baseUrl;

      const routeOf = (doc: DocMeta): string => {
        const rel = doc.permalink.startsWith(baseUrl)
          ? doc.permalink.slice(baseUrl.length)
          : doc.permalink.replace(/^\//, '');
        // category-index permalinks ('api/') keep a trailing slash even with
        // trailingSlash: false — the HTML still lands at api.html
        return rel.replace(/\/+$/, '');
      };
      // `trailingSlash: false` serves `<route>.html`; the twin sits at
      // `<route>.md`. The root route ('') becomes index.md.
      const twinRelPath = (doc: DocMeta): string =>
        `${routeOf(doc) || 'index'}.md`;
      const twinUrl = (doc: DocMeta): string => siteRoot + twinRelPath(doc);

      const rawSource = (doc: DocMeta): string =>
        fs.readFileSync(
          doc.source.replace(/^@site/, context.siteDir),
          'utf-8',
        );

      // 1 — a Markdown twin for every page, guides and API reference alike
      for (const doc of version.docs) {
        const twinPath = path.join(outDir, twinRelPath(doc));
        fs.mkdirSync(path.dirname(twinPath), { recursive: true });
        fs.writeFileSync(
          twinPath,
          normalizeBody(rawSource(doc), doc.title) + '\n',
        );
      }

      const docById = new Map(version.docs.map((doc) => [doc.id, doc]));
      const guideIds = flattenSidebarIds(
        version.sidebars.docs ?? [],
      ).filter((id) => !id.startsWith('api/'));
      const guides = guideIds.map((id) => {
        const doc = docById.get(id);
        if (!doc) throw new Error(`llms-txt: sidebar id without doc: ${id}`);
        return doc;
      });

      // The description in llms.txt is curated frontmatter, never an excerpt.
      const missing = guides.filter(
        (doc) =>
          typeof doc.frontMatter.description !== 'string' ||
          !doc.frontMatter.description.trim(),
      );
      if (missing.length) {
        throw new Error(
          `llms-txt: guide page(s) without a frontmatter description: ${missing
            .map((doc) => doc.source)
            .join(', ')}`,
        );
      }

      // 2 — llms.txt: flat index with one-line descriptions
      const apiIndex = docById.get('api/index');
      const llmsTxt = [
        `# ${siteConfig.title}`,
        '',
        `> ${siteConfig.tagline}. The links below point to the Markdown twin of each documentation page; every HTML page under ${siteRoot} is also served as Markdown at the same URL plus \`.md\`.`,
        '',
        '## Documentation',
        '',
        ...guides.map(
          (doc) =>
            `- [${doc.title}](${twinUrl(doc)}): ${doc.frontMatter.description}`,
        ),
        ...(apiIndex
          ? [
              '',
              '## API reference',
              '',
              `- [API index](${twinUrl(apiIndex)}): generated API reference (typedoc); every api/ page has a Markdown twin at its URL plus \`.md\``,
            ]
          : []),
        '',
        '## Optional',
        '',
        `- [llms-full.txt](${siteRoot}llms-full.txt): all documentation pages concatenated into one file`,
        `- [AI-INSTRUCTOR.md](https://raw.githubusercontent.com/singerla/pptx-automizer/main/AI-INSTRUCTOR.md): compact instructions for AI assistants writing code with pptx-automizer (also shipped in the npm package)`,
        '',
      ].join('\n');
      fs.writeFileSync(path.join(outDir, 'llms.txt'), llmsTxt);

      // 3 — llms-full.txt: the guide corpus in one file (the API reference
      // stays out — it is large, generated, and indexed via llms.txt)
      const sections = guides.map((doc) => {
        const [h1, ...rest] = normalizeBody(rawSource(doc), doc.title).split(
          '\n',
        );
        return [h1, '', `> Source: ${twinUrl(doc)}`, ...rest].join('\n');
      });
      const llmsFullTxt = [
        `# ${siteConfig.title} — full documentation`,
        '',
        `> ${siteConfig.tagline}. Every documentation page in one file; the per-page index is ${siteRoot}llms.txt.`,
        '',
        sections.join('\n\n---\n\n'),
        '',
      ].join('\n');
      fs.writeFileSync(path.join(outDir, 'llms-full.txt'), llmsFullTxt);
    },
  };
}
