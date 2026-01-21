import * as React from 'react';
import { DefaultButton, Icon, IconButton, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import type { ISuggestionComment, ISuggestionCommentState } from '../types';
import { isRichTextValueEmpty } from '../../utils/text';
import RichTextEditor from './RichTextEditor';

interface ICommentSectionProps {
  comment: ISuggestionCommentState;
  onToggle: () => void;
  onToggleComposer: () => void;
  onCommentDraftChange: (value: string) => void;
  onSubmitComment: () => void;
  onDeleteComment: (comment: ISuggestionComment) => void;
  onDeleteSuggestion: () => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
  canDeleteSuggestion: boolean;
}

const CommentSection: React.FC<ICommentSectionProps> = ({
  comment,
  onToggle,
  onToggleComposer,
  onCommentDraftChange,
  onSubmitComment,
  onDeleteComment,
  onDeleteSuggestion,
  formatDateTime,
  isLoading,
  canDeleteSuggestion
}) => {
  const isDraftEmpty: boolean = isRichTextValueEmpty(comment.draftText);
  const isSubmitDisabled: boolean = isDraftEmpty || comment.isSubmitting || isLoading;

  return (
    <div className={styles.commentSection}>
      <div className={styles.commentHeader}>
        <button
          type="button"
          id={comment.toggleId}
          className={styles.commentToggleButton}
          onClick={onToggle}
          aria-expanded={comment.isExpanded}
          aria-controls={comment.regionId}
        >
          <Icon
            iconName={comment.isExpanded ? 'ChevronDownSmall' : 'ChevronRightSmall'}
            className={styles.commentToggleIcon}
          />
          <span className={styles.commentHeading}>{strings.CommentsLabel}</span>
          <span className={styles.commentCount}>({comment.resolvedCount})</span>
        </button>
        {(comment.canAddComment || canDeleteSuggestion) && (
          <div className={styles.commentActions}>
            {comment.canAddComment && (
              <DefaultButton
                className={styles.commentAddButton}
                text={
                  comment.isComposerVisible
                    ? strings.HideCommentInputButtonText
                    : strings.AddCommentButtonText
                }
                onClick={onToggleComposer}
                disabled={isLoading || comment.isSubmitting}
              />
            )}
            {canDeleteSuggestion && (
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                className={styles.commentDeleteSuggestionButton}
                title={strings.RemoveSuggestionButtonLabel}
                ariaLabel={strings.RemoveSuggestionButtonLabel}
                onClick={onDeleteSuggestion}
                disabled={isLoading}
              />
            )}
          </div>
        )}
      </div>
      <div
        id={comment.regionId}
        role="region"
        aria-labelledby={comment.toggleId}
        className={`${styles.commentContent} ${comment.isExpanded ? '' : styles.commentContentCollapsed}`}
        hidden={!comment.isExpanded}
      >
        {comment.isExpanded && (
          comment.isLoading ? (
            <Spinner label={strings.LoadingCommentsLabel} size={SpinnerSize.small} />
          ) : !comment.hasLoaded ? null : (
            <>
              {comment.canAddComment && comment.isComposerVisible && (
                <div className={styles.commentComposer}>
                  <RichTextEditor
                    label={strings.CommentInputLabel}
                    value={comment.draftText}
                    onChange={(newValue) => onCommentDraftChange(newValue)}
                    placeholder={strings.CommentInputPlaceholder}
                    disabled={comment.isSubmitting || isLoading}
                  />
                  <PrimaryButton
                    className={styles.commentComposerSubmit}
                    text={strings.SubmitCommentButtonText}
                    onClick={onSubmitComment}
                    disabled={isSubmitDisabled}
                  />
                </div>
              )}
              {comment.comments.length === 0 ? (
                <p className={styles.commentEmpty}>{strings.NoCommentsLabel}</p>
              ) : (
                <ul className={styles.commentList}>
                  {comment.comments.map((commentItem) => (
                    <li key={commentItem.id} className={styles.commentItem}>
                      <div className={styles.commentItemTopRow}>
                        <div className={styles.commentMeta}>
                          <span className={styles.commentAuthor}>
                            {commentItem.author || strings.UnknownCommentAuthorLabel}
                          </span>
                          <span className={styles.commentTimestamp}>
                            {commentItem.createdDateTime
                              ? formatDateTime(commentItem.createdDateTime)
                              : strings.UnknownCommentDateLabel}
                          </span>
                        </div>
                        {comment.canDeleteComments ? (
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            className={styles.commentDeleteButton}
                            title={strings.DeleteCommentButtonLabel}
                            ariaLabel={strings.DeleteCommentButtonLabel}
                            onClick={() => onDeleteComment(commentItem)}
                            disabled={isLoading}
                          />
                        ) : (
                          <span className={styles.commentMetaPlaceholder} aria-hidden="true" />
                        )}
                      </div>
                      <div className={styles.commentText} dangerouslySetInnerHTML={{ __html: commentItem.text }} />
                    </li>
                  ))}
                </ul>
              )}
            </>
          )
        )}
      </div>
    </div>
  );
};

export default CommentSection;
