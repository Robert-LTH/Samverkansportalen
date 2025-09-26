import * as React from 'react';

import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import ImprovementService, {
  ISuggestionItem,
  SuggestionStatus
} from '../services/ImprovementService';

const STATUSES: SuggestionStatus[] = ['Föreslagen', 'Pågående', 'Genomförd', 'Avslutad'];

const STATUS_STYLES: Record<SuggestionStatus, string> = {
  Föreslagen: styles.statusProposed,
  Pågående: styles.statusOngoing,
  Genomförd: styles.statusCompleted,
  Avslutad: styles.statusArchived
};

const classNames = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === 'string') {
    return error;
  }

  return 'Ett oväntat fel inträffade.';
};

const formatDate = (value: string): string => {
  try {
    return new Date(value).toLocaleDateString('sv-SE', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch (error) {
    return value;
  }
};

const HelloWorld: React.FC<IHelloWorldProps> = ({ sp, currentUser }) => {
  const service = React.useMemo(() => new ImprovementService(sp), [sp]);
  const [suggestions, setSuggestions] = React.useState<ISuggestionItem[]>([]);
  const [remainingVotes, setRemainingVotes] = React.useState<number>(5);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [creating, setCreating] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>();
  const [title, setTitle] = React.useState<string>('');
  const [details, setDetails] = React.useState<string>('');
  const [searchTerm, setSearchTerm] = React.useState<string>('');
  const [busyIds, setBusyIds] = React.useState<Set<number>>(new Set());

  const refresh = React.useCallback(async () => {
    const data = await service.loadSuggestions(currentUser);
    setSuggestions(data.items);
    setRemainingVotes(data.remainingVotes);
  }, [service, currentUser]);

  React.useEffect(() => {
    let isMounted = true;
    const initialise = async (): Promise<void> => {
      setLoading(true);
      setError(undefined);
      try {
        await service.ensureInfrastructure();
        if (!isMounted) {
          return;
        }
        await refresh();
      } catch (err) {
        if (isMounted) {
          setError(getErrorMessage(err));
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    };

    initialise();

    return () => {
      isMounted = false;
    };
  }, [service, refresh]);

  const updateBusyState = (id: number, busy: boolean): void => {
    setBusyIds((current) => {
      const next = new Set(current);
      if (busy) {
        next.add(id);
      } else {
        next.delete(id);
      }
      return next;
    });
  };

  const handleCreateSuggestion = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const trimmedTitle = title.trim();
    const trimmedDetails = details.trim();

    if (!trimmedTitle || !trimmedDetails) {
      setError('Ange både titel och beskrivning för ditt förslag.');
      return;
    }

    setCreating(true);
    setError(undefined);

    try {
      await service.createSuggestion(trimmedTitle, trimmedDetails);
      setTitle('');
      setDetails('');
      await refresh();
    } catch (err) {
      setError(getErrorMessage(err));
    } finally {
      setCreating(false);
    }
  };

  const handleVoteToggle = async (suggestion: ISuggestionItem) => {
    if (busyIds.has(suggestion.id)) {
      return;
    }

    updateBusyState(suggestion.id, true);
    setError(undefined);

    try {
      if (suggestion.userHasVoted) {
        await service.removeVote(suggestion.id, currentUser);
      } else {
        if (remainingVotes <= 0) {
          setError('Du har inga röster kvar att använda.');
          return;
        }
        await service.castVote(suggestion.id, currentUser);
      }
      await refresh();
    } catch (err) {
      setError(getErrorMessage(err));
    } finally {
      updateBusyState(suggestion.id, false);
    }
  };

  const handleStatusChange = async (suggestion: ISuggestionItem, status: SuggestionStatus) => {
    if (busyIds.has(suggestion.id) || suggestion.status === status) {
      return;
    }

    updateBusyState(suggestion.id, true);
    setError(undefined);

    try {
      await service.updateStatus(suggestion.id, status);
      await refresh();
    } catch (err) {
      setError(getErrorMessage(err));
    } finally {
      updateBusyState(suggestion.id, false);
    }
  };

  const filteredSuggestions = React.useMemo(() => {
    const needle = searchTerm.trim().toLowerCase();
    if (!needle) {
      return suggestions;
    }

    return suggestions.filter((suggestion) => {
      const haystack = `${suggestion.title} ${suggestion.details} ${suggestion.author} ${suggestion.status}`.toLowerCase();
      return haystack.includes(needle);
    });
  }, [suggestions, searchTerm]);

  return (
    <section className={styles.portal} aria-busy={loading}>
      <header className={styles.header}>
        <div>
          <h1 className={styles.title}>Förbättringsportalen</h1>
          <p className={styles.subtitle}>
            Föreslå förbättringar, rösta på dina favoriter och följ upp hur arbetet fortskrider.
          </p>
        </div>
        <div className={styles.statusPanel}>
          <span className={styles.userName}>{currentUser.displayName}</span>
          <span className={styles.voteBadge}>
            Röster kvar: <strong>{remainingVotes}</strong>
          </span>
        </div>
      </header>

      {error && (
        <div className={styles.error} role="alert">
          {error}
        </div>
      )}

      <div className={styles.toolbar}>
        <input
          type="search"
          className={styles.search}
          placeholder="Sök efter titel, beskrivning eller status"
          value={searchTerm}
          onChange={(event) => setSearchTerm(event.target.value)}
          aria-label="Sök bland förslag"
        />
      </div>

      <form className={styles.createForm} onSubmit={handleCreateSuggestion}>
        <h2>Skapa nytt förbättringsförslag</h2>
        <label className={styles.field}>
          <span>Titel</span>
          <input
            type="text"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            placeholder="Beskriv förslaget kort"
            required
          />
        </label>

        <label className={styles.field}>
          <span>Beskrivning</span>
          <textarea
            value={details}
            onChange={(event) => setDetails(event.target.value)}
            placeholder="Förklara varför förslaget behövs och vad det innebär"
            rows={4}
            required
          />
        </label>

        <button type="submit" className={styles.primaryButton} disabled={creating}>
          {creating ? 'Sparar…' : 'Spara förslag'}
        </button>
      </form>

      <section aria-live="polite" className={styles.listSection}>
        <h2>Alla förslag</h2>
        {loading ? (
          <div className={styles.placeholder}>Hämtar förslag…</div>
        ) : filteredSuggestions.length === 0 ? (
          <div className={styles.placeholder}>Inga förslag matchar din sökning ännu.</div>
        ) : (
          <ul className={styles.suggestionList}>
            {filteredSuggestions.map((suggestion) => (
              <li key={suggestion.id} className={styles.suggestionCard}>
                <div className={styles.cardHeader}>
                  <div>
                    <h3>{suggestion.title}</h3>
                    <span className={styles.meta}>
                      Skapad {formatDate(suggestion.created)} av {suggestion.author}
                    </span>
                  </div>
                  <span
                    className={classNames(styles.statusBadge, STATUS_STYLES[suggestion.status])}
                  >
                    {suggestion.status}
                  </span>
                </div>

                <p className={styles.description}>{suggestion.details}</p>

                <div className={styles.cardFooter}>
                  <div className={styles.voteCount}>
                    Röster: <strong>{suggestion.voteCount}</strong>
                  </div>
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={classNames(
                        styles.secondaryButton,
                        busyIds.has(suggestion.id) ? styles.busyButton : undefined
                      )}
                      disabled={busyIds.has(suggestion.id) || (!suggestion.userHasVoted && remainingVotes <= 0)}
                      onClick={() => handleVoteToggle(suggestion)}
                    >
                      {suggestion.userHasVoted ? 'Återta röst' : 'Rösta'}
                    </button>

                    <label className={styles.statusSelectLabel}>
                      <span className="ms-screenReaderOnly">Uppdatera status</span>
                      <select
                        className={styles.statusSelect}
                        value={suggestion.status}
                        onChange={(event) =>
                          handleStatusChange(suggestion, event.target.value as SuggestionStatus)
                        }
                        disabled={busyIds.has(suggestion.id)}
                      >
                        {STATUSES.map((status) => (
                          <option key={status} value={status}>
                            {status}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                </div>
              </li>
            ))}
          </ul>
        )}
      </section>
    </section>
  );
};

export default HelloWorld;
