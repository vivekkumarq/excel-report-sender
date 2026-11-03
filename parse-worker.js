// One worksheet per message: inflate its compressed bytes and parse the XML, off the main thread.
import { inflateSync } from './vendor/fflate.js';
import { parseSheet } from './xlsx.js';

const dec = new TextDecoder();
let sst = [];
onmessage = ({ data }) => {
  if (data.sst) { sst = data.sst; return; }
  try {
    const raw = data.method ? inflateSync(data.bytes, { out: new Uint8Array(data.size) }) : data.bytes;
    postMessage({ id: data.id, result: parseSheet(dec.decode(raw), sst) });
  } catch (e) {
    postMessage({ id: data.id, error: String(e?.message || e) });
  }
};
