// Tiny Web Worker pool. Each job goes to the worker with the least pending work (cost given by the caller),
// so submitting the biggest jobs first gives "longest job first" scheduling. Workers reply { id, result | error }.
export function pool(url, { max = navigator.hardwareConcurrency || 4, init } = {}) {
  const workers = [], waiting = new Map();
  let seq = 0;
  const spawn = () => {
    const w = new Worker(url, { type: 'module' });
    w.load = 0;
    w.onmessage = ({ data }) => {
      const p = waiting.get(data.id);
      waiting.delete(data.id);
      w.load -= p.cost;
      data.error ? p.fail(new Error(data.error)) : p.ok(data.result);
    };
    w.onerror = e => {
      e.preventDefault();
      for (const [id, p] of waiting) if (p.w === w) { waiting.delete(id); p.fail(new Error('worker failed to start')); }
    };
    if (init) w.postMessage(init);
    workers.push(w);
    return w;
  };
  return {
    run(msg, { cost = 1, transfer = [] } = {}) {
      const w = workers.length < max ? spawn() : workers.reduce((a, b) => (b.load < a.load ? b : a));
      w.load += cost;
      return new Promise((ok, fail) => {
        const id = seq++;
        waiting.set(id, { ok, fail, cost, w });
        w.postMessage({ ...msg, id }, transfer);
      });
    },
    get size() { return workers.length; },
    close() { workers.forEach(w => w.terminate()); },
  };
}
