const STATUS_DROPDOWN_FONT_SIZE_REM: number = 0.65;
const STATUS_DROPDOWN_FONT_WEIGHT: string = '600';
const STATUS_DROPDOWN_LETTER_SPACING_EM: number = 0.05;
const STATUS_DROPDOWN_HORIZONTAL_PADDING_REM: number = 2.5;
const STATUS_DROPDOWN_CARET_PADDING_REM: number = 1.5;
const STATUS_DROPDOWN_LIST_PADDING_PX: number = 16;
const statusDropdownWidthCache: Map<string, number> = new Map();

export const measureStatusDropdownWidth = (values: string[]): number | undefined => {
  if (typeof document === 'undefined' || values.length === 0) {
    return undefined;
  }

  const cacheKey: string = values.join('|');
  const cachedWidth: number | undefined = statusDropdownWidthCache.get(cacheKey);
  if (cachedWidth) {
    return cachedWidth;
  }

  const body: HTMLElement | null = document.body;
  if (!body) {
    return undefined;
  }

  const span: HTMLSpanElement = document.createElement('span');
  span.style.position = 'absolute';
  span.style.visibility = 'hidden';
  span.style.whiteSpace = 'nowrap';
  span.style.fontSize = `${STATUS_DROPDOWN_FONT_SIZE_REM}rem`;
  span.style.fontWeight = STATUS_DROPDOWN_FONT_WEIGHT;
  span.style.letterSpacing = `${STATUS_DROPDOWN_LETTER_SPACING_EM}em`;
  span.style.textTransform = 'uppercase';
  span.style.fontFamily = window.getComputedStyle(body).fontFamily || 'Segoe UI';
  body.appendChild(span);

  let maxWidth: number = 0;
  values.forEach((value) => {
    span.textContent = value;
    maxWidth = Math.max(maxWidth, span.getBoundingClientRect().width);
  });

  body.removeChild(span);

  if (maxWidth <= 0) {
    return undefined;
  }

  const rootFontSize: number = parseFloat(
    window.getComputedStyle(document.documentElement).fontSize || '16'
  );
  const extraWidth: number =
    (STATUS_DROPDOWN_HORIZONTAL_PADDING_REM + STATUS_DROPDOWN_CARET_PADDING_REM) * rootFontSize +
    STATUS_DROPDOWN_LIST_PADDING_PX;
  const measuredWidth: number = Math.ceil(maxWidth + extraWidth);

  statusDropdownWidthCache.set(cacheKey, measuredWidth);
  return measuredWidth;
};
