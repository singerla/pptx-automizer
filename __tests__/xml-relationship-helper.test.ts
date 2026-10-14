import { DOMParser } from '@xmldom/xmldom';
import { XmlHelper, XmlRelationshipHelper } from '../src';
import { XmlDocument, XmlElement } from '../src/types/xml-types';
import IArchive from '../src/interfaces/iarchive';

/**
 * Unit coverage for the merged relationship parsing (#55): one
 * implementation behind `XmlRelationshipHelper.parseRelationTarget`, the
 * `XmlHelper.*` names delegate to it.
 */

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const relsDoc = (body: string): XmlDocument =>
  new DOMParser().parseFromString(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`,
    'application/xml',
  ) as unknown as XmlDocument;

const rel = (id: string, type: string, target: string): XmlElement =>
  relsDoc(
    `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`,
  ).getElementsByTagName('Relationship')[0];

describe('XmlRelationshipHelper.parseRelationTarget', () => {
  test('extracts file name parts and the trailing number', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId4', 'chart', '../charts/chart12.xml'),
    );

    expect(target.rId).toBe('rId4');
    expect(target.file).toBe('../charts/chart12.xml');
    expect(target.filename).toBe('chart12.xml');
    expect(target.filenameBase).toBe('chart');
    expect(target.filenameExt).toBe('xml');
    expect(target.number).toBe(12);
    expect(target.prefix).toBeUndefined();
    expect(target.subtype).toBeUndefined();
  });

  test('a file name without a number yields 0', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId1', 'slideLayout', '../slideLayouts/slideLayout.xml'),
    );
    expect(target.number).toBe(0);
    expect(target.filenameBase).toBe('slideLayout');
  });

  test('a matching prefix sets prefix, subtype and updateTargetIndex', () => {
    const element = rel('rId2', 'chart', '../charts/chart3.xml');
    const target = XmlRelationshipHelper.parseRelationTarget(
      element,
      '../charts/chart',
    );

    expect(target.prefix).toBe('../charts/chart');
    expect(target.subtype).toBe('chart');
    expect(target.number).toBe(3);

    target.updateTargetIndex(7);
    expect(element.getAttribute('Target')).toBe('../charts/chart7.xml');

    target.updateTargetValue('../charts/chart8.xml');
    expect(target.getTargetValue()).toBe('../charts/chart8.xml');

    target.updateId('rId99-created');
    expect(element.getAttribute('Id')).toBe('rId99-created');
  });

  test('a non-matching prefix returns the target without prefix', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId2', 'chart', '../charts/chart3.xml'),
      '../media/image',
    );

    expect(target).toBeDefined();
    expect(target.prefix).toBeUndefined();
    expect(target.subtype).toBe('image');
  });

  test('matchByPrefix drops non-matching targets', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId2', 'chart', '../charts/chart3.xml'),
      '../media/image',
      true,
    );
    expect(target).toBeUndefined();
  });

  test('the relationship type must match the prefix subtype', () => {
    // same target path, but an image relationship: not a chart
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId2', 'image', '../charts/chart3.xml'),
      '../charts/chart',
      true,
    );
    expect(target).toBeUndefined();
  });

  test('absolute /ppt/ targets (pptxgenjs) match a relative prefix', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId2', 'chart', '/ppt/charts/chart1.xml'),
      '../charts/chart',
      true,
    );
    expect(target).toBeDefined();
    expect(target.prefix).toBe('../charts/chart');
    expect(target.number).toBe(1);
  });

  test('../diagrams/data addresses diagramData relationships', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId5', 'diagramData', '../diagrams/data2.xml'),
      '../diagrams/data',
      true,
    );
    expect(target).toBeDefined();
    expect(target.subtype).toBe('diagramData');
    expect(target.number).toBe(2);
  });

  test('embedded packages match any prefix in their directory', () => {
    const target = XmlRelationshipHelper.parseRelationTarget(
      rel('rId1', 'package', '../embeddings/Microsoft_Excel_Worksheet1.xlsx'),
      '../embeddings/Microsoft_Excel_Worksheet',
      true,
    );
    expect(target).toBeDefined();
    expect(target.filenameExt).toBe('xlsx');
    expect(target.number).toBe(1);
  });

  test('returns undefined for a non-element', () => {
    expect(
      XmlRelationshipHelper.parseRelationTarget(undefined as XmlElement),
    ).toBeUndefined();
  });
});

describe('XmlRelationshipHelper rel ids', () => {
  const rels = relsDoc(
    `<Relationship Id="rId3" Type="${REL_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId10-created" Type="${REL_NS}/image" Target="../media/image1.png"/>` +
      `<Relationship Id="rId2" Type="${REL_NS}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>`,
  );

  test('getMaxId ignores the -created suffix', () => {
    const nodes = rels.documentElement.childNodes;
    expect(XmlRelationshipHelper.getMaxId(nodes, 'Id')).toBe(10);
    expect(XmlRelationshipHelper.getMaxId(nodes, 'Id', true)).toBe(11);
    expect(XmlRelationshipHelper.getMaxId(nodes, 'Id', true, 50)).toBe(50);
  });

  test('getNextRelId keeps the rId<max+1>-created format', async () => {
    const archive = {
      readXml: async () => rels,
    } as unknown as IArchive;

    expect(
      await XmlRelationshipHelper.getNextRelId(archive, 'any.rels'),
    ).toBe('rId11-created');
  });

  test('getTargetsByRelationshipType and ...ByPrefix read from an archive', async () => {
    const archive = {
      readXml: async () => rels,
    } as unknown as IArchive;

    const byType = await XmlRelationshipHelper.getTargetsByRelationshipType(
      archive,
      'any.rels',
      `${REL_NS}/image`,
    );
    expect(byType).toHaveLength(1);
    expect(byType[0].rId).toBe('rId10-created');
    expect(byType[0].file).toBe('../media/image1.png');

    const byPrefix =
      await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
        archive,
        'any.rels',
        ['../slideLayouts/slideLayout', '../notesSlides/notesSlide'],
      );
    expect(byPrefix.map((t) => t.rId)).toEqual(['rId3', 'rId2']);
    expect(byPrefix[0].subtype).toBe('slideLayout');
  });
});

describe('deprecated XmlHelper delegates', () => {
  test('XmlHelper.parseRelationTarget behaves like XmlRelationshipHelper', () => {
    const element = rel('rId2', 'chart', '../charts/chart3.xml');
    const viaXmlHelper = XmlHelper.parseRelationTarget(element, '../charts/chart');
    const direct = XmlRelationshipHelper.parseRelationTarget(
      element,
      '../charts/chart',
    );

    expect(viaXmlHelper.prefix).toBe(direct.prefix);
    expect(viaXmlHelper.subtype).toBe(direct.subtype);
    expect(viaXmlHelper.number).toBe(direct.number);
    expect(
      XmlHelper.parseRelationTarget(element, '../media/image').prefix,
    ).toBeUndefined();
  });

  test('XmlHelper.getMaxId delegates', () => {
    const nodes = relsDoc(
      `<Relationship Id="rId7" Type="${REL_NS}/slide" Target="slides/slide1.xml"/>`,
    ).documentElement.childNodes;
    expect(XmlHelper.getMaxId(nodes, 'Id', true)).toBe(8);
  });
});
