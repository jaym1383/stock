window.stockCache = {
  async access(mode, value) {
    const db = await new Promise((resolve, reject) => {
      const request = indexedDB.open('jaymart-stock-data', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('snapshots');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction('snapshots', mode);
        const store = tx.objectStore('snapshots');
        const request = mode === 'readonly' ? store.get('latest') : store.put(value, 'latest');
        tx.oncomplete = () => resolve(request.result);
        tx.onerror = tx.onabort = () => reject(tx.error);
      });
    } finally { db.close(); }
  },
  read() { return this.access('readonly'); },
  write(value) { return this.access('readwrite', value); }
};
