type Area = Map<string, unknown>;

function storageArea(data: Area) {
  return {
    async get(key?: string | string[] | null) {
      const keys = key == null ? [...data.keys()] : Array.isArray(key) ? key : [key];
      const out: Record<string, unknown> = {};
      for (const k of keys) if (data.has(k)) out[k] = structuredClone(data.get(k));
      return out;
    },
    async set(items: Record<string, unknown>) {
      for (const [k, v] of Object.entries(items)) data.set(k, structuredClone(v));
    },
    async remove(key: string | string[]) {
      for (const k of Array.isArray(key) ? key : [key]) data.delete(k);
    },
  };
}

/** Just enough of the chrome.* surface for the runner: two storage areas and host permissions. */
export function installFakeChrome() {
  const local: Area = new Map();
  const session: Area = new Map();
  const state = { permissionGranted: true };
  (globalThis as Record<string, unknown>).chrome = {
    storage: { local: storageArea(local), session: storageArea(session) },
    permissions: { contains: async () => state.permissionGranted },
    runtime: { id: 'test-extension-id' },
  };
  return { local, session, state };
}
