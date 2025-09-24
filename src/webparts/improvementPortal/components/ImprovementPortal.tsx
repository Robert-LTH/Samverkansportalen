import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { SPFI } from '@pnp/sp';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SearchBox,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField
} from 'office-ui-fabric-react';
import SuggestionService from '../services/SuggestionService';
import {
  ISuggestionWithVotes,
  SuggestionStatus,
  activeStatuses,
  suggestionStatusOptions
} from '../models/ImprovementModels';
import styles from './ImprovementPortal.module.scss';

const MAX_VOTES = 5;

export interface IImprovementPortalProps {
  description: string;
  sp: SPFI;
}

type FeedbackMessage = {
  text: string;
  type: MessageBarType;
};

const statusDropdownOptions: IDropdownOption[] = suggestionStatusOptions.map((option) => ({
  key: option.key,
  text: option.text
}));

const ImprovementPortal: React.FunctionComponent<IImprovementPortalProps> = (props: IImprovementPortalProps) => {
  const serviceRef = useRef<SuggestionService>();
  if (!serviceRef.current) {
    serviceRef.current = new SuggestionService(props.sp);
  }
  const service = serviceRef.current;

  const [suggestions, setSuggestions] = useState<ISuggestionWithVotes[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [working, setWorking] = useState<boolean>(false);
  const [error, setError] = useState<string | undefined>();
  const [feedback, setFeedback] = useState<FeedbackMessage | undefined>();
  const [searchValue, setSearchValue] = useState<string>('');
  const [currentUserId, setCurrentUserId] = useState<number | undefined>();
  const [newTitle, setNewTitle] = useState<string>('');
  const [newDescription, setNewDescription] = useState<string>('');
  const [titleError, setTitleError] = useState<string | undefined>();

  const statusLabelMap = useMemo(() => {
    const map = new Map<SuggestionStatus, string>();
    suggestionStatusOptions.forEach((option) => map.set(option.key, option.text));
    return map;
  }, []);

  const loadInitialData = useCallback(async () => {
    setLoading(true);
    setError(undefined);
    try {
      await service.ensureSetup();
      const user = await service.getCurrentUser();
      const userId = user.id;
      if (userId === undefined) {
        throw new Error('Kunde inte läsa in användarinformation.');
      }
      setCurrentUserId(userId);
      const loadedSuggestions = await service.getSuggestions('', userId);
      setSuggestions(loadedSuggestions);
      setSearchValue('');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Det gick inte att läsa in data.';
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [service]);

  useEffect(() => {
    loadInitialData();
  }, [loadInitialData]);

  const refreshSuggestions = useCallback(
    async (query?: string) => {
      if (currentUserId === undefined) {
        return;
      }
      setLoading(true);
      setError(undefined);
      try {
        const items = await service.getSuggestions(query !== undefined ? query : searchValue, currentUserId);
        setSuggestions(items);
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Kunde inte uppdatera listan med förslag.';
        setError(message);
      } finally {
        setLoading(false);
      }
    },
    [service, currentUserId, searchValue]
  );

  const availableVotes = useMemo(() => {
    const activeVotes = suggestions.filter((item) => item.userHasActiveVote).length;
    const remaining = MAX_VOTES - activeVotes;
    return remaining > 0 ? remaining : 0;
  }, [suggestions]);

  const handleSearch = useCallback(
    async (value?: string) => {
      const query = value ? value.trim() : '';
      setSearchValue(query);
      await refreshSuggestions(query);
    },
    [refreshSuggestions]
  );

  const handleVote = useCallback(
    async (suggestion: ISuggestionWithVotes) => {
      const userId = currentUserId;
      if (userId === undefined) {
        return;
      }
      if (availableVotes <= 0) {
        setFeedback({ text: 'Du har inga röster kvar att använda just nu.', type: MessageBarType.warning });
        return;
      }
      try {
        setWorking(true);
        await service.addVote(suggestion.id, userId);
        await refreshSuggestions();
        setFeedback({ text: `Du har röstat på "${suggestion.title}".`, type: MessageBarType.success });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Kunde inte registrera din röst.';
        setFeedback({ text: message, type: MessageBarType.error });
      } finally {
        setWorking(false);
      }
    },
    [service, currentUserId, availableVotes, refreshSuggestions]
  );

  const handleWithdrawVote = useCallback(
    async (suggestion: ISuggestionWithVotes) => {
      if (!suggestion.userVoteId) {
        return;
      }
      try {
        setWorking(true);
        await service.withdrawVote(suggestion.userVoteId);
        await refreshSuggestions();
        setFeedback({ text: 'Din röst har tagits bort från förslaget.', type: MessageBarType.success });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Kunde inte ta bort din röst.';
        setFeedback({ text: message, type: MessageBarType.error });
      } finally {
        setWorking(false);
      }
    },
    [service, refreshSuggestions]
  );

  const handleStatusChange = useCallback(
    async (suggestion: ISuggestionWithVotes, status: SuggestionStatus) => {
      if (suggestion.status === status) {
        return;
      }
      try {
        setWorking(true);
        await service.updateSuggestionStatus(suggestion.id, status);
        await refreshSuggestions();
        const statusLabel = statusLabelMap.get(status) || status;
        setFeedback({ text: `Förslaget markerades som ${statusLabel}.`, type: MessageBarType.success });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Kunde inte uppdatera status för förslaget.';
        setFeedback({ text: message, type: MessageBarType.error });
      } finally {
        setWorking(false);
      }
    },
    [service, refreshSuggestions, statusLabelMap]
  );

  const handleCreateSuggestion = useCallback(
    async (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      setFeedback(undefined);
      if (!newTitle || newTitle.trim().length === 0) {
        setTitleError('Ange en titel för förslaget.');
        return;
      }
      try {
        setWorking(true);
        await service.createSuggestion(newTitle.trim(), newDescription.trim());
        setNewTitle('');
        setNewDescription('');
        setTitleError(undefined);
        await refreshSuggestions();
        setFeedback({ text: 'Förslaget har lagts till.', type: MessageBarType.success });
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Kunde inte spara förslaget.';
        setFeedback({ text: message, type: MessageBarType.error });
      } finally {
        setWorking(false);
      }
    },
    [service, newTitle, newDescription, refreshSuggestions]
  );

  const resetForm = useCallback(() => {
    setNewTitle('');
    setNewDescription('');
    setTitleError(undefined);
  }, []);

  const activeSuggestions = useMemo(
    () =>
      suggestions
        .filter((item) => activeStatuses.includes(item.status))
        .sort((a, b) => {
          if (b.activeVotes !== a.activeVotes) {
            return b.activeVotes - a.activeVotes;
          }
          return new Date(b.created).getTime() - new Date(a.created).getTime();
        }),
    [suggestions]
  );

  const archivedSuggestions = useMemo(
    () =>
      suggestions
        .filter((item) => !activeStatuses.includes(item.status))
        .sort((a, b) => new Date(b.created).getTime() - new Date(a.created).getTime()),
    [suggestions]
  );

  const renderSuggestion = (suggestion: ISuggestionWithVotes) => {
    const createdDate = suggestion.created ? new Date(suggestion.created).toLocaleDateString() : '';
    const isActive = activeStatuses.includes(suggestion.status);
    const statusLabel = statusLabelMap.get(suggestion.status) || suggestion.status;

    return (
      <div key={suggestion.id} className={styles.suggestionCard}>
        <div className={styles.cardHeader}>
          <Text variant="large">{suggestion.title}</Text>
          <span className={styles.statusBadge}>{statusLabel}</span>
        </div>
        <div className={styles.meta}>
          {suggestion.createdBy?.title && (
            <span>
              Skapad av {suggestion.createdBy.title} den {createdDate}
            </span>
          )}
        </div>
        {suggestion.description && <div className={styles.description}>{suggestion.description}</div>}
        <div className={styles.voteActions}>
          <span className={styles.voteCounter}>
            {isActive ? `Aktiva röster: ${suggestion.activeVotes}` : `Historiska röster: ${suggestion.totalVotes}`}
          </span>
          {isActive && suggestion.userHasActiveVote && <span>Du har röstat på detta förslag.</span>}
          {isActive && !suggestion.userHasActiveVote && availableVotes <= 0 && (
            <span>Du har inga röster kvar att använda.</span>
          )}
          {isActive && suggestion.userHasActiveVote ? (
            <DefaultButton
              text="Återta röst"
              onClick={() => handleWithdrawVote(suggestion)}
              disabled={working}
            />
          ) : null}
          {isActive && !suggestion.userHasActiveVote ? (
            <PrimaryButton text="Rösta" onClick={() => handleVote(suggestion)} disabled={working || availableVotes <= 0} />
          ) : null}
          {!isActive && suggestion.userHasAnyVote && (
            <span>Din röst har återlämnats.</span>
          )}
        </div>
        <div className={styles.statusControls}>
          <span>Status:</span>
          <Dropdown
            options={statusDropdownOptions}
            selectedKey={suggestion.status}
            onChange={(_, option) => option && handleStatusChange(suggestion, option.key as SuggestionStatus)}
            disabled={working}
          />
        </div>
      </div>
    );
  };

  return (
    <div className={styles.portal}>
      <div className={styles.header}>
        <Text variant="xLarge">Förbättringsportalen</Text>
        {props.description && <Text>{props.description}</Text>}
      </div>

      {feedback && (
        <MessageBar
          className={styles.messageBar}
          messageBarType={feedback.type}
          isMultiline={false}
          onDismiss={() => setFeedback(undefined)}
          dismissButtonAriaLabel="Stäng"
        >
          {feedback.text}
        </MessageBar>
      )}

      {error && (
        <MessageBar
          className={styles.messageBar}
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setError(undefined)}
        >
          {error}
        </MessageBar>
      )}

      <form className={styles.form} onSubmit={handleCreateSuggestion}>
        <Text variant="large">Lägg till nytt förslag</Text>
        <TextField
          label="Titel"
          value={newTitle}
          onChange={(_, value) => {
            setNewTitle(value || '');
            setTitleError(undefined);
          }}
          required
          errorMessage={titleError}
        />
        <TextField
          label="Beskrivning"
          multiline
          rows={4}
          value={newDescription}
          onChange={(_, value) => setNewDescription(value || '')}
        />
        <div className={styles.formButtons}>
          <PrimaryButton type="submit" text="Spara förslag" disabled={working} />
          <DefaultButton text="Rensa" onClick={resetForm} disabled={working} />
        </div>
      </form>

      <div className={styles.controls}>
        <div className={styles.searchRow}>
          <SearchBox
            placeholder="Sök efter titel eller beskrivning"
            value={searchValue}
            onSearch={handleSearch}
            onChange={(_, value) => setSearchValue(value || '')}
            onClear={() => handleSearch('')}
          />
          <Text variant="mediumPlus">Tillgängliga röster: {availableVotes} av {MAX_VOTES}</Text>
        </div>
      </div>

      {loading ? (
        <Spinner size={SpinnerSize.large} label="Laddar förslag..." />
      ) : suggestions.length === 0 ? (
        <div className={styles.emptyState}>Inga förslag hittades. Lägg till ett nytt eller ändra sökningen.</div>
      ) : (
        <Stack tokens={{ childrenGap: 12 }}>
          {activeSuggestions.length > 0 && <Text className={styles.sectionTitle}>Aktiva förbättringsförslag</Text>}
          {activeSuggestions.map((suggestion) => renderSuggestion(suggestion))}

          {archivedSuggestions.length > 0 && <Text className={styles.sectionTitle}>Avslutade eller borttagna förslag</Text>}
          {archivedSuggestions.map((suggestion) => renderSuggestion(suggestion))}
        </Stack>
      )}
    </div>
  );
};

export default ImprovementPortal;
