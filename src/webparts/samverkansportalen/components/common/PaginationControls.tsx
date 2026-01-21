import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';

interface IPaginationControlsProps {
  page: number;
  hasPrevious: boolean;
  hasNext: boolean;
  totalPages?: number;
  onPrevious: () => void;
  onNext: () => void;
}

const PaginationControls: React.FC<IPaginationControlsProps> = ({
  page,
  hasPrevious,
  hasNext,
  totalPages,
  onPrevious,
  onNext
}) => {
  if (!hasPrevious && !hasNext && page <= 1) {
    return null;
  }

  const normalizedTotalPages: number | undefined =
    typeof totalPages === 'number' && Number.isFinite(totalPages)
      ? Math.max(1, Math.floor(totalPages))
      : undefined;
  const label: string = normalizedTotalPages
    ? strings.PaginationPageCountLabel
        .replace('{0}', page.toString())
        .replace('{1}', normalizedTotalPages.toString())
    : strings.PaginationPageLabel.replace('{0}', page.toString());

  return (
    <div className={styles.paginationControls}>
      <DefaultButton text={strings.PreviousButtonText} onClick={onPrevious} disabled={!hasPrevious} />
      <span className={styles.paginationInfo} aria-live="polite">
        {label}
      </span>
      <DefaultButton text={strings.NextButtonText} onClick={onNext} disabled={!hasNext} />
    </div>
  );
};

export default PaginationControls;
