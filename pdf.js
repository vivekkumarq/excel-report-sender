// Smallest useful PDF writer: JPEG pictures placed on pages, embedded as-is (DCTDecode), so nothing is re-encoded.
// pages: [{ pw, ph, images: [{ jpeg: Uint8Array, w, h, x, y, dw, dh }] }]
//   pw/ph page size in points; w/h image pixels; x/y/dw/dh placement in points from the page's top-left.
export function jpegPdf(pages) {
  const enc = new TextEncoder(), parts = [], offsets = [], kids = [];
  let size = 0, next = 3;
  const put = p => { const b = typeof p === 'string' ? enc.encode(p) : p; parts.push(b); size += b.length; };
  const obj = (n, dict, stream) => {
    offsets[n] = size;
    put(`${n} 0 obj\n${dict}\n`);
    if (stream) { put('stream\n'); put(stream); put('\nendstream\n'); }
    put('endobj\n');
  };
  const f = v => +v.toFixed(2);
  put('%PDF-1.4\n%âãÏÓ\n');
  const body = [];
  for (const p of pages) {
    const page = next++, content = next++, imgs = p.images.map(() => next++);
    kids.push(`${page} 0 R`);
    const draw = p.images.map((im, i) => `q ${f(im.dw)} 0 0 ${f(im.dh)} ${f(im.x)} ${f(p.ph - im.y - im.dh)} cm /Im${i} Do Q`).join('\n');
    const xobj = imgs.map((n, i) => `/Im${i} ${n} 0 R`).join(' ');
    body.push(() => {
      obj(page, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${f(p.pw)} ${f(p.ph)}] /Resources << /XObject << ${xobj} >> >> /Contents ${content} 0 R >>`);
      obj(content, `<< /Length ${draw.length} >>`, draw);
      p.images.forEach((im, i) => obj(imgs[i], `<< /Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${im.jpeg.length} >>`, im.jpeg));
    });
  }
  obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
  obj(2, `<< /Type /Pages /Kids [${kids.join(' ')}] /Count ${pages.length} >>`);
  body.forEach(write => write());
  const xref = size;
  put(`xref\n0 ${next}\n0000000000 65535 f \n`);
  for (let n = 1; n < next; n++) put(`${String(offsets[n]).padStart(10, '0')} 00000 n \n`);
  put(`trailer\n<< /Size ${next} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  return new Blob(parts, { type: 'application/pdf' });
}
