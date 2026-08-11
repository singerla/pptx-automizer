import { Chart } from '../shapes/chart';
import { Diagram } from '../shapes/diagram';
import { Hyperlink } from '../shapes/hyperlink';
import { Image } from '../shapes/image';
import { OLEObject } from '../shapes/ole';
import type HasShapes from './has-shapes';

/**
 * Copies all content related to a slide/master/layout — charts, images,
 * diagrams, OLE objects and hyperlinks — from the source archive into the
 * target archive when the slide is appended.
 *
 * Extracted from HasShapes (ROADMAP Phase 2).
 */
export class RelatedContentCopier {
  constructor(private host: HasShapes) {}

  async copy(): Promise<void> {
    await this.copyCharts();
    await this.copyImages();
    await this.copyDiagrams();
    await this.copyOleObjects();
    await this.copyHyperlinks();
  }

  private async copyCharts(): Promise<void> {
    const host = this.host;
    const charts = await Chart.getAllOnSlide(host.sourceArchive, host.relsPath);
    for (const chart of charts) {
      await new Chart(
        {
          mode: 'append',
          target: chart,
          sourceArchive: host.sourceArchive,
          sourceSlideNumber: host.sourceNumber,
        },
        host.targetType,
      ).modifyOnAddedSlide(host.targetTemplate, host.targetNumber);
    }
  }

  private async copyImages(): Promise<void> {
    const host = this.host;
    const images = await Image.getAllOnSlide(host.sourceArchive, host.relsPath);
    for (const image of images) {
      await new Image(
        {
          mode: 'append',
          target: image,
          sourceArchive: host.sourceArchive,
          sourceSlideNumber: host.sourceNumber,
        },
        host.targetType,
      ).modifyOnAddedSlide(host.targetTemplate, host.targetNumber);
    }
  }

  private async copyDiagrams(): Promise<void> {
    const host = this.host;
    const diagrams = await Diagram.getAllOnSlide(
      host.sourceArchive,
      host.relsPath,
    );
    for (const diagram of diagrams) {
      await new Diagram(
        {
          mode: 'append',
          target: diagram,
          sourceArchive: host.sourceArchive,
          sourceSlideNumber: host.sourceNumber,
        },
        host.targetType,
      ).modifyOnAddedSlide(host.targetTemplate, host.targetNumber);
    }
  }

  private async copyOleObjects(): Promise<void> {
    const host = this.host;
    const oleObjects = await OLEObject.getAllOnSlide(
      host.sourceArchive,
      host.relsPath,
    );
    for (const oleObject of oleObjects) {
      await new OLEObject(
        {
          mode: 'append',
          target: oleObject,
          sourceArchive: host.sourceArchive,
          sourceSlideNumber: host.sourceNumber,
        },
        host.targetType,
        host.sourceArchive,
      ).modifyOnAddedSlide(host.targetTemplate, host.targetNumber, oleObjects);
    }
  }

  private async copyHyperlinks(): Promise<void> {
    const host = this.host;
    const hyperlinks = await Hyperlink.getAllOnSlide(
      host.sourceArchive,
      host.relsPath,
    );
    for (const hyperlink of hyperlinks) {
      const hyperlinkInstance = new Hyperlink(
        {
          mode: 'append',
          target: hyperlink,
          sourceArchive: host.sourceArchive,
          sourceSlideNumber: host.sourceNumber,
          sourceRid: hyperlink.rId,
        },
        host.targetType,
        host.sourceArchive,
        hyperlink.isExternal ? 'external' : 'internal',
        hyperlink.file,
      );

      // Ensure the target property is properly set
      hyperlinkInstance.target = hyperlink;

      await hyperlinkInstance.modifyOnAddedSlide(
        host.targetTemplate,
        host.targetNumber,
      );
    }
  }
}
