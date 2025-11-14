type SrcRect = { l?: number; t?: number; r?: number; b?: number }; // thousandths-of-percent
const K = 100000;

// Convert the source rectangle to fractions of the image width and height for easier calculation
function toFractions(src: SrcRect): { l: number; t: number; r: number; b: number } {
  return {
    l: (src.l ?? 0) / K,
    t: (src.t ?? 0) / K,
    r: (src.r ?? 0) / K,
    b: (src.b ?? 0) / K,
  };
}

// Infer the container aspect ratio from the original image dimensions and the current source rectangle
export function inferContainerAr(
  oldImageWidth: number,
  oldImageHeight: number,
  currentSrcRect: SrcRect // read from the slide; missing attrs mean 0
): number {
  const { l, t, r, b } = toFractions(currentSrcRect);
  const wf = 1 - (l + r);
  const hf = 1 - (t + b);
  const oldAr = oldImageWidth / oldImageHeight;
  return oldAr * (wf / hf || 1); // guard vs hf=0
}

// Compute the new source rectangle for a new image based on the container aspect ratio and the new image dimensions
export function computeSrcRectForNewImage(
  containerAr: number,
  newImageWidth: number,
  newImageHeight: number
): SrcRect {
  const newAr = newImageWidth / newImageHeight;
  if (!isFinite(containerAr) || !isFinite(newAr) || containerAr <= 0 || newAr <= 0) {
    return { l: 0, t: 0, r: 0, b: 0 };
  }

  if (newAr > containerAr) {
    // new image is wider than the container -> crop width
    const visibleW = containerAr / newAr;
    const crop = Math.max(0, Math.min(0.5, (1 - visibleW) / 2));
    const cropHorizontal = Math.round(crop * K);
    return { l: cropHorizontal, r: cropHorizontal, t: 0, b: 0 };
  } else if (newAr < containerAr) {
    // new image is taller than the container -> crop height
    const visibleH = newAr / containerAr;
    const crop = Math.max(0, Math.min(0.5, (1 - visibleH) / 2));
    const cropVertical = Math.round(crop * K);
    return { l: 0, r: 0, t: cropVertical, b: cropVertical };
  } else {
    // equal AR -> no crop needed
    return { l: 0, t: 0, r: 0, b: 0 };
  }
}