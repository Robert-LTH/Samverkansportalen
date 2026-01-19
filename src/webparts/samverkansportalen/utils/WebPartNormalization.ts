import { type IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import * as strings from 'SamverkansportalenWebPartStrings';
import {
  DEFAULT_COMMENTS_LIST_SUFFIX,
  DEFAULT_STATUS_DEFINITIONS,
  DEFAULT_SUGGESTIONS_LIST_TITLE,
  DEFAULT_TOTAL_VOTES_PER_USER,
  DEFAULT_VOTES_LIST_SUFFIX
} from '../components/ISamverkansportalenProps';
import type { IGraphStatusItem } from '../services/GraphSuggestionsService';
import type { IStatusDefinition } from '../SamverkansportalenWebPart.types';

export const normalizeListTitle = (value?: string): string => {
  const trimmed: string = (value ?? '').trim();
  return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
};

export const getDefaultVoteListTitle = (listTitle: string): string => {
  const trimmed: string = listTitle.trim();
  if (trimmed.length === 0) {
    return strings.DefaultVotesListTitle;
  }

  return `${trimmed}${DEFAULT_VOTES_LIST_SUFFIX}`;
};

export const normalizeVoteListTitle = (value?: string, listTitle?: string): string => {
  const trimmed: string = (value ?? '').trim();
  const normalizedListTitle: string = normalizeListTitle(listTitle);
  return trimmed.length > 0 ? trimmed : getDefaultVoteListTitle(normalizedListTitle);
};

export const getDefaultCommentListTitle = (listTitle: string): string => {
  const trimmed: string = listTitle.trim();
  if (trimmed.length === 0) {
    return strings.DefaultCommentsListTitle;
  }

  return `${trimmed}${DEFAULT_COMMENTS_LIST_SUFFIX}`;
};

export const normalizeCommentListTitle = (value?: string, listTitle?: string): string => {
  const trimmed: string = (value ?? '').trim();
  const normalizedListTitle: string = normalizeListTitle(listTitle);
  return trimmed.length > 0 ? trimmed : getDefaultCommentListTitle(normalizedListTitle);
};

export const normalizeOptionalListTitle = (value?: string): string | undefined => {
  const trimmed: string = (value ?? '').trim();
  return trimmed.length > 0 ? trimmed : undefined;
};

export const normalizeHeaderText = (value: string | undefined, fallback: string): string => {
  const trimmed: string = (value ?? '').trim();
  return trimmed.length > 0 ? trimmed : fallback;
};

export const parseStatusDefinitions = (value?: string): string[] => {
  const source: string =
    typeof value === 'string' && value.trim().length > 0 ? value : DEFAULT_STATUS_DEFINITIONS;
  const segments: string[] = source.split(/[\n,;]/);
  const seen: Set<string> = new Set();
  const results: string[] = [];

  segments.forEach((segment) => {
    const trimmed: string = segment.trim();

    if (!trimmed) {
      return;
    }

    const key: string = trimmed.toLowerCase();

    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    results.push(trimmed);
  });

  if (results.length === 0 && source !== DEFAULT_STATUS_DEFINITIONS) {
    return parseStatusDefinitions(DEFAULT_STATUS_DEFINITIONS);
  }

  return results.length > 0 ? results : ['Active', 'Done'];
};

export const parseStatusSortOrder = (value: unknown): number | undefined => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === 'string') {
    const parsed: number = parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return undefined;
};

export const parseBooleanField = (value: unknown): boolean => {
  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    const normalized: string = value.trim().toLowerCase();
    return normalized === 'true' || normalized === '1' || normalized === 'yes';
  }

  return false;
};

export const mapStatusDefinitions = (items: IGraphStatusItem[]): IStatusDefinition[] =>
  items
    .map((item) => {
      const rawTitle: unknown = item.fields?.Title;

      if (typeof rawTitle !== 'string') {
        return undefined;
      }

      const title: string = rawTitle.trim();

      if (!title) {
        return undefined;
      }

      const order: number | undefined = parseStatusSortOrder(item.fields?.SortOrder);
      const isCompleted: boolean = parseBooleanField(item.fields?.IsCompleted);

      return {
        id: item.id,
        title,
        order,
        isCompleted
      } as IStatusDefinition;
    })
    .filter((definition): definition is IStatusDefinition => typeof definition !== 'undefined');

export const normalizeStatusDefinitions = (value?: string): string => {
  const statuses: string[] = parseStatusDefinitions(value);
  return statuses.join('\n');
};

export const normalizeCompletedStatus = (
  value: string | undefined,
  statuses: string[]
): string => {
  if (statuses.length === 0) {
    return 'Done';
  }

  const trimmed: string = (value ?? '').trim();

  if (trimmed.length > 0) {
    const match: string | undefined = statuses.find(
      (status) => status.toLowerCase() === trimmed.toLowerCase()
    );

    if (match) {
      return match;
    }
  }

  return statuses[statuses.length - 1];
};

export const normalizeDeniedStatus = (
  value: string | undefined,
  statuses: string[],
  completedStatus: string
): string | undefined => {
  const trimmed: string = (value ?? '').trim();

  if (!trimmed) {
    return undefined;
  }

  const match: string | undefined = statuses.find(
    (status) => status.toLowerCase() === trimmed.toLowerCase()
  );

  return match ?? trimmed;
};

export const normalizeDefaultStatus = (
  value: string | undefined,
  statuses: string[],
  completedStatus: string
): string => {
  if (statuses.length === 0) {
    return completedStatus;
  }

  const trimmed: string = (value ?? '').trim();

  if (trimmed.length > 0) {
    const match: string | undefined = statuses.find(
      (status) => status.toLowerCase() === trimmed.toLowerCase()
    );

    if (match) {
      return match;
    }
  }

  const firstActive: string | undefined = statuses.find(
    (status) => status.toLowerCase() !== completedStatus.toLowerCase()
  );

  return firstActive ?? completedStatus;
};

export const normalizeTotalVotesPerUser = (value: string | undefined): number => {
  const trimmed: string = (value ?? '').trim();

  if (!trimmed) {
    return DEFAULT_TOTAL_VOTES_PER_USER;
  }

  const parsed: number = Number(trimmed);

  if (!Number.isFinite(parsed)) {
    return DEFAULT_TOTAL_VOTES_PER_USER;
  }

  const rounded: number = Math.floor(parsed);
  return rounded > 0 ? rounded : DEFAULT_TOTAL_VOTES_PER_USER;
};

export const validateTotalVotesPerUser = (value: string): string => {
  const trimmed: string = (value ?? '').trim();

  if (!trimmed) {
    return '';
  }

  const parsed: number = Number(trimmed);

  if (!Number.isFinite(parsed) || parsed <= 0 || !Number.isInteger(parsed)) {
    return strings.TotalVotesPerUserFieldErrorMessage;
  }

  return '';
};

export const getDropdownOptionText = (option: IPropertyPaneDropdownOption): string => {
  const text: string | undefined = typeof option.text === 'string' ? option.text.trim() : undefined;

  if (text && text.length > 0) {
    return text;
  }

  const key: unknown = option.key;

  if (typeof key === 'string' || typeof key === 'number' || typeof key === 'boolean') {
    const normalizedKey: string = String(key).trim();

    if (normalizedKey.length > 0) {
      return normalizedKey;
    }
  }

  return '';
};
