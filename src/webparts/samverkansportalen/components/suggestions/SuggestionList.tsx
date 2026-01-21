import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import type {
  CommentAction,
  ISuggestionItem,
  ISuggestionViewModel,
  SuggestionAction
} from '../types';
import CommentSection from '../common/CommentSection';
import SuggestionStatusControl from '../common/SuggestionStatusControl';
import SuggestionTimestamps from '../common/SuggestionTimestamps';

interface ISuggestionCardsProps {
  viewModels: ISuggestionViewModel[];
  onToggleVote: SuggestionAction;
  onChangeStatus: (item: ISuggestionItem, status: string) => void;
  onDeleteSuggestion: SuggestionAction;
  onSubmitComment: SuggestionAction;
  onCommentDraftChange: (item: ISuggestionItem, value: string) => void;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  onToggleCommentComposer: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
  statuses: string[];
}

const SuggestionCards: React.FC<ISuggestionCardsProps> = ({
  viewModels,
  onToggleVote,
  onChangeStatus,
  onDeleteSuggestion,
  onSubmitComment,
  onCommentDraftChange,
  onDeleteComment,
  onToggleComments,
  onToggleCommentComposer,
  formatDateTime,
  isLoading,
  statuses
}) => (
  <ul className={styles.suggestionList}>
    {viewModels.map(({ item, interaction, comment }) => (
      <li key={item.id} className={styles.suggestionCard}>
        <div className={styles.cardHeader}>
          <div className={styles.cardText}>
            <div className={styles.cardMeta}>
              <span
                className={styles.entryId}
                aria-label={strings.EntryAriaLabel.replace('{0}', item.id.toString())}
              >
                #{item.id}
              </span>
              <span className={styles.categoryBadge}>{item.category}</span>
              {item.subcategory && <span className={styles.subcategoryBadge}>{item.subcategory}</span>}
              <SuggestionStatusControl
                statuses={statuses}
                value={item.status}
                isEditable={interaction.canAdvanceSuggestionStatus}
                isDisabled={isLoading}
                onChange={(status) => onChangeStatus(item, status)}
              />
            </div>
            <h4 className={styles.suggestionTitle}>{item.title}</h4>
            <SuggestionTimestamps item={item} formatDateTime={formatDateTime} />
            {item.description && (
              <div
                className={styles.suggestionDescription}
                dangerouslySetInnerHTML={{ __html: item.description }}
              />
            )}
          </div>
          <div
            className={styles.voteBadge}
            aria-label={`${item.votes} ${item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}`}
          >
            <span className={styles.voteNumber}>{item.votes}</span>
            <span className={styles.voteText}>
              {item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}
            </span>
            <div className={styles.voteActions}>
              {interaction.isVotingAllowed ? (
                <PrimaryButton
                  text={interaction.hasVoted ? strings.RemoveVoteButtonText : strings.VoteButtonText}
                  onClick={() => onToggleVote(item)}
                  disabled={interaction.disableVote}
                />
              ) : (
                <DefaultButton text={strings.VotesClosedText} disabled />
              )}
            </div>
          </div>
        </div>
        <CommentSection
          comment={comment}
          onToggle={() => onToggleComments(item.id)}
          onToggleComposer={() => onToggleCommentComposer(item.id)}
          onCommentDraftChange={(value) => onCommentDraftChange(item, value)}
          onSubmitComment={() => onSubmitComment(item)}
          onDeleteComment={(commentItem) => onDeleteComment(item, commentItem)}
          onDeleteSuggestion={() => onDeleteSuggestion(item)}
          formatDateTime={formatDateTime}
          isLoading={isLoading}
          canDeleteSuggestion={interaction.canDeleteSuggestion}
        />
      </li>
    ))}
  </ul>
);

interface ISuggestionTableProps {
  viewModels: ISuggestionViewModel[];
  onToggleVote: SuggestionAction;
  onChangeStatus: (item: ISuggestionItem, status: string) => void;
  onDeleteSuggestion: SuggestionAction;
  onSubmitComment: SuggestionAction;
  onCommentDraftChange: (item: ISuggestionItem, value: string) => void;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  onToggleCommentComposer: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
  statuses: string[];
  showMetadataInIdColumn: boolean;
}

const SuggestionTable: React.FC<ISuggestionTableProps> = ({
  viewModels,
  onToggleVote,
  onChangeStatus,
  onDeleteSuggestion,
  onSubmitComment,
  onCommentDraftChange,
  onDeleteComment,
  onToggleComments,
  onToggleCommentComposer,
  formatDateTime,
  isLoading,
  statuses,
  showMetadataInIdColumn
}) => (
  <div className={styles.tableWrapper}>
    <table className={styles.suggestionTable}>
      <thead>
        <tr>
          <th scope="col" className={styles.tableHeaderId}>
            {strings.SuggestionTableEntryColumnLabel}
          </th>
          <th scope="col" className={styles.tableHeaderSuggestion}>
            {strings.SuggestionTableSuggestionColumnLabel}
          </th>
          {!showMetadataInIdColumn && (
            <>
              <th scope="col" className={styles.tableHeaderCategory}>
                {strings.CategoryLabel}
              </th>
              <th scope="col" className={styles.tableHeaderSubcategory}>
                {strings.SubcategoryLabel}
              </th>
              <th scope="col" className={styles.tableHeaderStatus}>
                {strings.StatusLabel}
              </th>
            </>
          )}
          <th scope="col" className={styles.tableHeaderVotes}>
            {strings.VotesAndActionsLabel}
          </th>
        </tr>
      </thead>
      <tbody>
        {viewModels.map(({ item, interaction, comment }) => (
          <React.Fragment key={item.id}>
            <tr className={styles.suggestionRow}>
              <td
                className={`${styles.tableCellId} ${showMetadataInIdColumn ? styles.tableCellIdWithMeta : ''}`}
                data-label={strings.SuggestionTableEntryColumnLabel}
              >
                <div className={styles.entryMetaColumn}>
                  <span
                    className={styles.entryId}
                    aria-label={strings.EntryAriaLabel.replace('{0}', item.id.toString())}
                  >
                    #{item.id}
                  </span>
                  {showMetadataInIdColumn && (
                    <div className={styles.entryMetaDetails}>
                      <span className={styles.categoryBadge}>{item.category}</span>
                      {item.subcategory ? (
                        <span className={styles.subcategoryBadge}>{item.subcategory}</span>
                      ) : (
                        <span className={styles.subcategoryPlaceholder}>
                          {strings.NoSubcategoriesAvailablePlaceholder}
                        </span>
                      )}
                      <div className={styles.inlineStatusControl}>
                        <SuggestionStatusControl
                          statuses={statuses}
                          value={item.status}
                          isEditable={interaction.canAdvanceSuggestionStatus}
                          isDisabled={isLoading}
                          onChange={(status) => onChangeStatus(item, status)}
                        />
                      </div>
                    </div>
                  )}
                </div>
              </td>
              <td
                className={styles.tableCellSuggestion}
                data-label={strings.SuggestionTableSuggestionColumnLabel}
              >
                <h4 className={styles.suggestionTitle}>{item.title}</h4>
                <SuggestionTimestamps item={item} formatDateTime={formatDateTime} />
                {item.description && (
                  <div
                    className={styles.suggestionDescription}
                    dangerouslySetInnerHTML={{ __html: item.description }}
                  />
                )}
              </td>
              {!showMetadataInIdColumn && (
                <>
                  <td className={styles.tableCellCategory} data-label={strings.CategoryLabel}>
                    <span className={styles.categoryBadge}>{item.category}</span>
                  </td>
                  <td className={styles.tableCellSubcategory} data-label={strings.SubcategoryLabel}>
                    {item.subcategory ? (
                      <span className={styles.subcategoryBadge}>{item.subcategory}</span>
                    ) : (
                      <span className={styles.subcategoryPlaceholder}>-</span>
                    )}
                  </td>
                  <td className={styles.tableCellStatus} data-label={strings.StatusLabel}>
                    <SuggestionStatusControl
                      statuses={statuses}
                      value={item.status}
                      isEditable={interaction.canAdvanceSuggestionStatus}
                      isDisabled={isLoading}
                      onChange={(status) => onChangeStatus(item, status)}
                    />
                  </td>
                </>
              )}
              <td className={styles.tableCellVotes} data-label={strings.VotesAndActionsLabel}>
                <div
                  className={styles.voteBadge}
                  aria-label={`${item.votes} ${item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}`}
                >
                  <span className={styles.voteNumber}>{item.votes}</span>
                  <span className={styles.voteText}>
                    {item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}
                  </span>
                  <div className={styles.voteActions}>
                    {interaction.isVotingAllowed ? (
                      <PrimaryButton
                        text={interaction.hasVoted ? '-' : '+'}
                        ariaLabel={
                          interaction.hasVoted ? strings.RemoveVoteButtonText : strings.VoteButtonText
                        }
                        onClick={() => onToggleVote(item)}
                        disabled={interaction.disableVote}
                      />
                    ) : (
                      <DefaultButton text={strings.VotesClosedText} disabled />
                    )}
                  </div>
                </div>
              </td>
            </tr>
            <tr className={styles.metaRow}>
              <td
                className={styles.metaCell}
                colSpan={showMetadataInIdColumn ? 3 : 6}
                data-label={strings.SuggestionTableDetailsColumnLabel}
              >
                <div className={styles.metaContent}>
                  <CommentSection
                    comment={comment}
                    onToggle={() => onToggleComments(item.id)}
                    onToggleComposer={() => onToggleCommentComposer(item.id)}
                    onCommentDraftChange={(value) => onCommentDraftChange(item, value)}
                    onSubmitComment={() => onSubmitComment(item)}
                    onDeleteComment={(commentItem) => onDeleteComment(item, commentItem)}
                    onDeleteSuggestion={() => onDeleteSuggestion(item)}
                    formatDateTime={formatDateTime}
                    isLoading={isLoading}
                    canDeleteSuggestion={interaction.canDeleteSuggestion}
                  />
                </div>
              </td>
            </tr>
          </React.Fragment>
        ))}
      </tbody>
    </table>
  </div>
);

interface ISuggestionListProps {
  viewModels: ISuggestionViewModel[];
  useTableLayout: boolean;
  showMetadataInIdColumn: boolean;
  isLoading: boolean;
  onToggleVote: SuggestionAction;
  onChangeStatus: (item: ISuggestionItem, status: string) => void;
  onDeleteSuggestion: SuggestionAction;
  onSubmitComment: SuggestionAction;
  onCommentDraftChange: (item: ISuggestionItem, value: string) => void;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  onToggleCommentComposer: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  statuses: string[];
}

const SuggestionList: React.FC<ISuggestionListProps> = ({
  viewModels,
  useTableLayout,
  showMetadataInIdColumn,
  isLoading,
  onToggleVote,
  onChangeStatus,
  onDeleteSuggestion,
  onSubmitComment,
  onCommentDraftChange,
  onDeleteComment,
  onToggleComments,
  onToggleCommentComposer,
  formatDateTime,
  statuses
}) => {
  if (viewModels.length === 0) {
    return (
      <div className={styles.suggestionListWrapper}>
        <p className={styles.emptyState}>{strings.NoSuggestionsLabel}</p>
      </div>
    );
  }

  const listContent: JSX.Element = useTableLayout ? (
    <SuggestionTable
      viewModels={viewModels}
      onToggleVote={onToggleVote}
      onChangeStatus={onChangeStatus}
      onDeleteSuggestion={onDeleteSuggestion}
      onSubmitComment={onSubmitComment}
      onCommentDraftChange={onCommentDraftChange}
      onDeleteComment={onDeleteComment}
      onToggleComments={onToggleComments}
      onToggleCommentComposer={onToggleCommentComposer}
      formatDateTime={formatDateTime}
      isLoading={isLoading}
      statuses={statuses}
      showMetadataInIdColumn={showMetadataInIdColumn}
    />
  ) : (
    <SuggestionCards
      viewModels={viewModels}
      onToggleVote={onToggleVote}
      onChangeStatus={onChangeStatus}
      onDeleteSuggestion={onDeleteSuggestion}
      onSubmitComment={onSubmitComment}
      onCommentDraftChange={onCommentDraftChange}
      onDeleteComment={onDeleteComment}
      onToggleComments={onToggleComments}
      onToggleCommentComposer={onToggleCommentComposer}
      formatDateTime={formatDateTime}
      isLoading={isLoading}
      statuses={statuses}
    />
  );

  return <div className={styles.suggestionListWrapper}>{listContent}</div>;
};

export default SuggestionList;
