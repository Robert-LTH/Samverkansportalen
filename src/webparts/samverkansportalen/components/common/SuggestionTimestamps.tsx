import * as React from 'react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import type { ISuggestionItem } from '../types';

interface ISuggestionTimestampsProps {
  item: ISuggestionItem;
  formatDateTime: (value: string) => string;
}

const SuggestionTimestamps: React.FC<ISuggestionTimestampsProps> = ({ item, formatDateTime }) => {
  const entries: { label: string; value: string }[] = [];
  const { createdDateTime, lastModifiedDateTime, completedDateTime, createdByLoginName } = item;

  if (createdDateTime) {
    entries.push({ label: strings.CreatedLabel, value: createdDateTime });
  }

  const shouldShowLastModified: boolean = !!lastModifiedDateTime && !completedDateTime;

  if (shouldShowLastModified && lastModifiedDateTime) {
    entries.push({ label: strings.LastModifiedLabel, value: lastModifiedDateTime });
  }

  if (completedDateTime) {
    entries.push({ label: strings.CompletedLabel, value: completedDateTime });
  }

  if (entries.length === 0) {
    return null;
  }

  return (
    <div className={styles.metadataSegment}>
      <span className={styles.authorRow}>
        <span className={styles.timestampLabel}>{strings.CreatedByLabel}:</span>
        <span className={styles.timestampValue}>{createdByLoginName}</span>
      </span>
      <span className={styles.timestampRow}>
        {entries.map((entry) => (
          <span key={entry.label} className={styles.timestampEntryEnd}>
            <span className={styles.timestampLabel}>{entry.label}:</span>
            <span className={styles.timestampValue}>{formatDateTime(entry.value)}</span>
          </span>
        ))}
      </span>
    </div>
  );
};

export default SuggestionTimestamps;
