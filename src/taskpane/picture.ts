/* Turning the drawn chart into a picture for inserting into the sheet. */

/* global document, XMLSerializer, Image, SVGSVGElement */

/** Rasterizes an SVG element to a Base64 PNG (no data: prefix), as Excel's addImage expects. */
export function svgToPngBase64(
  svg: SVGSVGElement,
  width: number,
  height: number,
  scale: number
): Promise<string> {
  const xml = new XMLSerializer().serializeToString(svg);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
  return new Promise((resolve, reject) => {
    const fail = () => reject(new Error("Couldn't create the chart image."));
    const img = new Image();
    img.onload = () => {
      // An error thrown in here wouldn't reject the promise by itself (e.g. a
      // canvas the browser treats as "tainted"), leaving Insert stuck for good.
      try {
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(width * scale);
        canvas.height = Math.round(height * scale);
        const ctx = canvas.getContext("2d");
        if (!ctx) return fail();
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, width, height);
        const png = canvas.toDataURL("image/png").split(",")[1];
        if (!png) return fail();
        resolve(png);
      } catch {
        fail();
      }
    };
    img.onerror = fail;
    img.src = url;
  });
}
