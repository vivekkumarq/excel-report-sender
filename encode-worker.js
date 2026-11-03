// Encodes one rendered report (an ImageBitmap) to PNG/JPEG on its own thread.
onmessage = async ({ data }) => {
  try {
    const { bitmap, type, quality } = data, canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    canvas.getContext('2d', { willReadFrequently: true }).drawImage(bitmap, 0, 0); // CPU canvas: no GPU read-back
    bitmap.close();
    postMessage({ id: data.id, result: await canvas.convertToBlob({ type, quality }) });
  } catch (e) {
    postMessage({ id: data.id, error: String(e?.message || e) });
  }
};
