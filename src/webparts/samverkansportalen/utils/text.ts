export const getPlainTextFromHtml = (value: string | undefined): string => {
  if (!value) {
    return '';
  }

  return value
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
};

export const isRichTextValueEmpty = (value: string): boolean => getPlainTextFromHtml(value).length === 0;
