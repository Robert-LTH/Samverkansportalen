import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import type {
  CommentAction,
  ISimilarSuggestionsQuery,
  ISuggestionItem,
  ISuggestionViewModel,
  SuggestionAction
} from '../types';
import PaginationControls from '../common/PaginationControls';
import SuggestionList from './SuggestionList';

interface ISimilarSuggestionsProps {
  viewModels: ISuggestionViewModel[];
  isLoading: boolean;
  query: ISimilarSuggestionsQuery;
  page: number;
  hasPrevious: boolean;
  hasNext: boolean;
  onPrevious: () => void;
  onNext: () => void;
  onToggleVote: SuggestionAction;
  onChangeStatus: (item: ISuggestionItem, status: string) => void;
  onDeleteSuggestion: SuggestionAction;
  onSubmitComment: SuggestionAction;
  onCommentDraftChange: (item: ISuggestionItem, value: string) => void;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  onToggleCommentComposer: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isProcessing: boolean;
  statuses: string[];
}

const SimilarSuggestions: React.FC<ISimilarSuggestionsProps> = ({
  viewModels,
  isLoading,
  query,
  page,
  hasPrevious,
  hasNext,
  onPrevious,
  onNext,
  onToggleVote,
  onChangeStatus,
  onDeleteSuggestion,
  onSubmitComment,
  onCommentDraftChange,
  onDeleteComment,
  onToggleComments,
  onToggleCommentComposer,
  formatDateTime,
  isProcessing,
  statuses
}) => {
  const hasTitleQuery: boolean = query.title.length > 0;
  const hasDescriptionQuery: boolean = query.description.length > 0;

  if (!hasTitleQuery && !hasDescriptionQuery) {
    return null;
  }

  const querySegments: { key: string; content: React.ReactNode }[] = [];

  if (hasTitleQuery) {
    querySegments.push({
      key: 'title',
      content: (
        <>
          {strings.SimilarSuggestionsQueryTitleLabel}{' '}
          <span className={styles.similarSuggestionsQueryValue}>&quot;{query.title}&quot;</span>
        </>
      )
    });
  }

  if (hasDescriptionQuery) {
    querySegments.push({
      key: 'description',
      content: (
        <>
          {strings.SimilarSuggestionsQueryDescriptionLabel}{' '}
          <span className={styles.similarSuggestionsQueryValue}>&quot;{query.description}&quot;</span>
        </>
      )
    });
  }

  const hasResults: boolean = viewModels.length > 0;

  return (
    <div className={styles.similarSuggestions} aria-live="polite">
      <div className={styles.similarSuggestionsHeader}>
        <h4 className={styles.similarSuggestionsTitle}>{strings.SimilarSuggestionsTitle}</h4>
        {!isLoading && hasResults && (
          <span className={styles.similarSuggestionsSummary}>
            {viewModels.length === 1
              ? strings.SingleMatchingSuggestionLabel
              : strings.MultipleMatchingSuggestionsLabel.replace('{0}', viewModels.length.toString())}
          </span>
        )}
      </div>
      <p className={styles.similarSuggestionsQuery}>
        {strings.SimilarSuggestionsQueryPrefix}{' '}
        {querySegments.map((segment, index) => (
          <React.Fragment key={segment.key}>
            {index > 0 && (
              <>
                {' '}
                {strings.SimilarSuggestionsQuerySeparator}
                {' '}
              </>
            )}
            {segment.content}
          </React.Fragment>
        ))}
      </p>
      {isLoading ? (
        <Spinner label={strings.SimilarSuggestionsLoadingLabel} size={SpinnerSize.small} />
      ) : hasResults ? (
        <>
          <div className={styles.similarSuggestionsResults}>
            <SuggestionList
              viewModels={viewModels}
              useTableLayout={false}
              showMetadataInIdColumn={false}
              isLoading={isProcessing}
              onToggleVote={onToggleVote}
              onChangeStatus={onChangeStatus}
              onDeleteSuggestion={onDeleteSuggestion}
              onSubmitComment={onSubmitComment}
              onCommentDraftChange={onCommentDraftChange}
              onDeleteComment={onDeleteComment}
              onToggleComments={onToggleComments}
              onToggleCommentComposer={onToggleCommentComposer}
              formatDateTime={formatDateTime}
              statuses={statuses}
            />
          </div>
          <PaginationControls
            page={page}
            hasPrevious={hasPrevious}
            hasNext={hasNext}
            onPrevious={onPrevious}
            onNext={onNext}
          />
        </>
      ) : (
        <p className={styles.noSimilarSuggestions}>{strings.NoSimilarSuggestionsLabel}</p>
      )}
    </div>
  );
};

export default SimilarSuggestions;
