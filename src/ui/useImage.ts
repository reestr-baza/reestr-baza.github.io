import { useEffect, useState } from 'react';
import { loadImageUrl, peekImageUrl } from '../storage/images';

/** URL фото из базы; пока грузится — undefined, если фото нет — null. */
export function useImageUrl(id: string | undefined, variant: 'thumb' | 'full' = 'thumb'): string | null | undefined {
  const [url, setUrl] = useState<string | null | undefined>(() => (id ? peekImageUrl(id, variant) : null));
  useEffect(() => {
    if (!id) {
      setUrl(null);
      return;
    }
    const have = peekImageUrl(id, variant);
    if (have) {
      setUrl(have);
      return;
    }
    let alive = true;
    setUrl(undefined);
    loadImageUrl(id, variant).then((u) => {
      if (alive) setUrl(u);
    });
    return () => {
      alive = false;
    };
  }, [id, variant]);
  return url;
}
