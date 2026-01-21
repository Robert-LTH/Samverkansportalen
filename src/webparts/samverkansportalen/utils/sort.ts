export const isSortDiagnosticsEnabled = (): boolean => {
  if (typeof window === 'undefined') {
    return false;
  }

  try {
    return new URLSearchParams(window.location.search).has('debugSort');
  } catch {
    return false;
  }
};

export const isClientSortForced = (): boolean => {
  if (typeof window === 'undefined') {
    return false;
  }

  try {
    return new URLSearchParams(window.location.search).has('forceSort');
  } catch {
    return false;
  }
};

export const getSortableDateValue = (value?: string): number => {
  if (!value) {
    return 0;
  }

  const parsed: number = Date.parse(value);
  return Number.isNaN(parsed) ? 0 : parsed;
};
