import { useSyncExternalStore } from 'react';
import type { Store } from '../model/store';
import type { Persister, SaveState } from '../storage/persist';

/** Единственные экземпляры хранилища и записи на диск, создаются в main.tsx. */
export let store: Store;
export let persister: Persister;

export function setInstances(s: Store, p: Persister) {
  store = s;
  persister = p;
}

/** Перерисовать компонент при любом изменении книги. */
export function useStoreVersion(): number {
  return useSyncExternalStore(store.subscribe, store.getVersion);
}

let saveState: SaveState = { status: 'saved', at: Date.now() };
const saveSubs = new Set<() => void>();

export function wireSaveState(p: Persister) {
  saveState = p.state;
  p.onSaveState((s) => {
    saveState = s;
    saveSubs.forEach((f) => f());
  });
}

export function useSaveState(): SaveState {
  return useSyncExternalStore(
    (f) => {
      saveSubs.add(f);
      return () => saveSubs.delete(f);
    },
    () => saveState,
  );
}
