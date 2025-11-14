import { XmlElement } from '../types/xml-types';
import { ImageStyle } from '../types/modify-types';
import slugify from 'slugify';
import {imageSize } from 'image-size';
import fs from 'fs';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { computeSrcRectForNewImage, inferContainerAr } from './compute-src-rect';

export default class ModifyImageHelper {
  /**
   * Update the "Target" attribute of a created image relation.
   * This will change the image itself. Load images with Automizer.loadMedia
   * @param filename name of target image in root template media folder.
   */
  static setRelationTarget = (filename: string) => {
    return (element: XmlElement, arg1: XmlElement): void => {
      arg1.setAttribute('Target', '../media/' + slugify(filename));
    };
  };

  /**
   * Update the "Target" attribute of a created image relation.
   * This will change the image itself. Load images with Automizer.loadMedia
   * This will also auto-crop the image to the new width and height,
   * based on the container aspect ratio, derived from the original image
   * and using the new image width and height based on the files loaded into
   * the presentation media folder using .loadMedia()
   * @param filename name of target image in root template media folder.
   * @param pres presentation properties (to access the root template archive)
   * @param newImageWidth width of the new image
   * @param newImageHeight height of the new image
   */
  static setRelationTargetCover = (
    filename: string,
    pres: IPresentationProps,
  ) => {
    return async (element: XmlElement, arg1: XmlElement): Promise<void> => {

      const newTarget = '../media/' + slugify(filename);
      const originalTarget = arg1.getAttribute('Target');
      const originalTargetPath = originalTarget.replace('../', 'ppt/');
      const originalImageDimensions = { width: 100, height: 100 };
      const newImageDimensions = { width: 100, height: 100 };


      // Get the new image dimensions, using the rootTemplate mediafiles array,
      // since we have it loaded into the presentation media folder using .loadMedia()
      // If we don't find the media file, we warn, but continue
      try {
        const mediaFile = pres.rootTemplate.mediaFiles.find(file => file.file === filename);
        if(!mediaFile) {
          throw new Error("Media file not found in template archive in path: " + filename);
        }
        const buffer = fs.readFileSync(mediaFile.filepath);
        const _dimensions = imageSize(buffer);
        newImageDimensions.width = _dimensions.width;
        newImageDimensions.height = _dimensions.height;
      } catch (error) {
        console.warn("Couldn't find media file in template archive in path.");
      }
      
      // Find the original image dimensions using the original target path from the original slide
      // using the rootTemplate archive file system and get the image dimensions.
      // If we don't find the original image, we warn, but continue
      // If we find the original image, we get the image dimensions and use this to reverse calculate
      // the aspect ratio and then use the new image dimensions to calculate and set the new crop on srcRect.
      // This results in the image being cropped in the image container to match the aspect ratio of the new image.
      try {
        if(pres.rootTemplate.archive.fileExists(originalTargetPath)) {
          const originalImage = await pres.rootTemplate.archive.read(originalTargetPath, "nodebuffer");
          const _dimensions = imageSize(originalImage);
          originalImageDimensions.width = _dimensions.width;
          originalImageDimensions.height = _dimensions.height;
        } else {
          throw new Error("Original image not found from template archive in path: " + originalTargetPath);
        }
  
        const srcRect = element.getElementsByTagName('a:srcRect')[0];
        const srcRectLeft = srcRect.getAttribute('l');
        const srcRectTop = srcRect.getAttribute('t');
        const srcRectRight = srcRect.getAttribute('r');
        const srcRectBottom = srcRect.getAttribute('b');
  
        const currentSrcRect = {
          l: srcRectLeft ? Number(srcRectLeft) : 0,
          t: srcRectTop ? Number(srcRectTop) : 0,
          r: srcRectRight ? Number(srcRectRight) : 0,
          b: srcRectBottom ? Number(srcRectBottom) : 0,
        }
  
        const containerAr = inferContainerAr(originalImageDimensions.width, originalImageDimensions.height, currentSrcRect);
        
        const newSrcRect = computeSrcRectForNewImage(containerAr, newImageDimensions.width, newImageDimensions.height);
        
        srcRect.setAttribute('l', String(newSrcRect.l));
        srcRect.setAttribute('t', String(newSrcRect.t));
        srcRect.setAttribute('r', String(newSrcRect.r));
        srcRect.setAttribute('b', String(newSrcRect.b));
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.warn("Skipped setting relation target cropped due to an error: " + errorMessage);
      }
      
      arg1.setAttribute('Target', newTarget);

    };
  }

  /*
    Update an existing duotone image overlay element (WIP)
    Apply a duotone color to an image p:blipFill -> a:blip fill element.
    Works best on white icons, see __tests__/media/feather.png
   */
  static setDuotoneFill =
    (duotoneParams: ImageStyle['duotone']) =>
    (element: XmlElement): void => {
      const blipFill = element.getElementsByTagName('p:blipFill');
      if (!blipFill) {
        return;
      }
      const duotone = blipFill.item(0).getElementsByTagName('a:duotone')[0];
      if (duotone) {
        if (duotoneParams?.color) {
          const srgbClr = duotone.getElementsByTagName('a:srgbClr')[0];
          if (srgbClr) {
            // Only sRgb supported
            srgbClr.setAttribute('val', String(duotoneParams.color.value));

            if (duotoneParams?.tint !== undefined) {
              // tint needs to be 0 - 100000
              const tint = srgbClr.getElementsByTagName('a:tint')[0];
              if (tint) {
                tint.setAttribute('val', String(duotoneParams.tint));
              }
            }

            if (duotoneParams?.satMod !== undefined) {
              const satMod = srgbClr.getElementsByTagName('a:satMod')[0];
              if (satMod) {
                satMod.setAttribute('val', String(duotoneParams.satMod));
              }
            }
          }
        }
        if (duotoneParams?.prstClr) {
          const prstClr = duotone.getElementsByTagName('a:prstClr')[0];
          if (prstClr) {
            // Only tested with "black" and "white"
            prstClr.setAttribute('val', String(duotoneParams.prstClr));
          }
        }
      }
    };
}
