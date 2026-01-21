/* eslint-disable max-lines */
import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TextField,
  Dropdown,
  type IDropdownOption
} from '@fluentui/react';
import { debounce } from '@microsoft/sp-lodash-subset';
import styles from './Samverkansportalen.module.scss';
import {
  DEFAULT_SUGGESTIONS_LIST_TITLE,
  DEFAULT_VOTES_LIST_SUFFIX,
  DEFAULT_COMMENTS_LIST_SUFFIX,
  DEFAULT_TOTAL_VOTES_PER_USER,
  type ISamverkansportalenProps
} from './ISamverkansportalenProps';
import TabHeader from './tabs/TabHeader';
import TabPanel from './tabs/TabPanel';
import RichTextEditor from './common/RichTextEditor';
import SuggestionSection from './suggestions/SuggestionSection';
import SuggestionList from './suggestions/SuggestionList';
import SimilarSuggestions from './suggestions/SimilarSuggestions';
import {
  type IFilterState,
  type IPaginatedSuggestionsState,
  type ISamverkansportalenState,
  type ISimilarSuggestionsQuery,
  type ISuggestionComment,
  type ISuggestionInteractionState,
  type ISuggestionItem,
  type ISuggestionViewModel,
  type ISubcategoryDefinition,
  type IVoteEntry,
  type MainTabKey
} from './types';
import { getPlainTextFromHtml, isRichTextValueEmpty } from '../utils/text';
import { getSortableDateValue, isClientSortForced, isSortDiagnosticsEnabled } from '../utils/sort';
import {
  type SuggestionCategory,
  type IGraphSuggestionItem,
  type IGraphSuggestionItemFields,
  type IGraphVoteItem,
  type IGraphSubcategoryItem,
  type IGraphCategoryItem,
  type IGraphCommentItem
} from '../services/GraphSuggestionsService';
import * as strings from 'SamverkansportalenWebPartStrings';

const DEFAULT_MAX_VOTES_PER_CATEGORY: number = DEFAULT_TOTAL_VOTES_PER_USER;
const FALLBACK_CATEGORIES: SuggestionCategory[] = [
  strings.DefaultCategoryChangeRequest,
  strings.DefaultCategoryWebinar,
  strings.DefaultCategoryArticle
];
const DEFAULT_SUGGESTION_CATEGORY: SuggestionCategory = FALLBACK_CATEGORIES[0];
const ALL_CATEGORY_FILTER_KEY: string = '__all_categories__';
const ALL_SUBCATEGORY_FILTER_KEY: string = '__all_subcategories__';
const ALL_STATUS_FILTER_KEY: string = '__all_statuses__';
const DEFAULT_SUGGESTIONS_PAGE_SIZE: number = 5;
const SUGGESTION_PAGE_SIZE_OPTIONS: number[] = [5, 10, 20];
const ADMIN_TOP_SUGGESTIONS_COUNT: number = 10;
const SIMILAR_SUGGESTIONS_DEBOUNCE_MS: number = 500;
const LIST_SEARCH_DEBOUNCE_MS: number = 300;
const MIN_SIMILAR_SUGGESTION_QUERY_LENGTH: number = 3;
const MAX_SIMILAR_SUGGESTIONS: number = 5;
const EMPTY_SIMILAR_SUGGESTIONS_QUERY: ISimilarSuggestionsQuery = { title: '', description: '' };
const SCROLL_CONTAINER_CLASS: string = 'c_WA0gy_hHQBj';

export default class Samverkansportalen extends React.Component<ISamverkansportalenProps, ISamverkansportalenState> {
  private _isMounted: boolean = false;
  private _currentListId?: string;
  private _currentVotesListId?: string;
  private _currentCommentsListId?: string;
  private _currentSubcategoryListId?: string;
  private _currentCategoryListId?: string;
  private _currentStatusListId?: string;
  private readonly _sectionIds: {
    add: { title: string; content: string };
    active: { title: string; content: string };
    completed: { title: string; content: string };
  };
  private readonly _commentSectionPrefix: string;
  private readonly _debouncedSimilarSuggestionsSearch: ReturnType<typeof debounce>;
  private readonly _debouncedActiveFilterSearch: ReturnType<typeof debounce>;
  private readonly _debouncedCompletedFilterSearch: ReturnType<typeof debounce>;
  private _pendingSimilarSuggestionsQuery?: ISimilarSuggestionsQuery;
  private readonly _rootRef: React.RefObject<HTMLElement>;
  private readonly _headerRef: React.RefObject<HTMLElement>;
  private _stickyUpdateFrame?: number;
  private _pivotLinksElement?: HTMLElement;
  private _isStickyActive: boolean = false;
  private _scrollContainer?: HTMLElement | Window;

  public constructor(props: ISamverkansportalenProps) {
    super(props);

    const uniquePrefix: string = `samverkansportalen-${Math.random().toString(36).slice(2, 10)}`;
    const { statuses, completedStatus, deniedStatus, defaultStatus } =
      this._deriveStatusStateFromProps(props);
    this._sectionIds = {
      add: { title: `${uniquePrefix}-add-title`, content: `${uniquePrefix}-add-content` },
      active: { title: `${uniquePrefix}-active-title`, content: `${uniquePrefix}-active-content` },
      completed: {
        title: `${uniquePrefix}-completed-title`,
        content: `${uniquePrefix}-completed-content`
      }
    };
    this._commentSectionPrefix = `${uniquePrefix}-comment`;
    this._debouncedSimilarSuggestionsSearch = debounce((query: ISimilarSuggestionsQuery) => {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._searchSimilarSuggestions(query);
    }, SIMILAR_SUGGESTIONS_DEBOUNCE_MS);
    this._debouncedActiveFilterSearch = debounce((filter: IFilterState) => {
      this._applyActiveFilter(filter);
    }, LIST_SEARCH_DEBOUNCE_MS);
    this._debouncedCompletedFilterSearch = debounce((filter: IFilterState) => {
      this._applyCompletedFilter(filter);
    }, LIST_SEARCH_DEBOUNCE_MS);
    this._rootRef = React.createRef<HTMLElement>();
    this._headerRef = React.createRef<HTMLElement>();

    this.state = {
      activeSuggestions: { items: [], page: 1, currentToken: undefined, nextToken: undefined, previousTokens: [] },
      completedSuggestions: { items: [], page: 1, currentToken: undefined, nextToken: undefined, previousTokens: [] },
      activePageSize: DEFAULT_SUGGESTIONS_PAGE_SIZE,
      completedPageSize: DEFAULT_SUGGESTIONS_PAGE_SIZE,
      activeSuggestionsTotal: undefined,
      completedSuggestionsTotal: undefined,
      isLoading: false,
      isActiveSuggestionsLoading: false,
      isCompletedSuggestionsLoading: false,
      newTitle: '',
      newDescription: '',
      newCategory: DEFAULT_SUGGESTION_CATEGORY,
      newSubcategoryKey: undefined,
      subcategories: [],
      categories: [...FALLBACK_CATEGORIES],
      statuses,
      completedStatus,
      deniedStatus,
      defaultStatus,
      availableVotesByCategory: {},
      isUnlimitedVotes: props.isCurrentUserAdmin,
      activeFilter: {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId: undefined,
        status: undefined
      },
      completedFilter: {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId: undefined,
        status: completedStatus,
        includeDenied: false
      },
      adminFilter: {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId: undefined,
        status: undefined
      },
      similarSuggestions: {
        items: [],
        page: 1,
        currentToken: undefined,
        nextToken: undefined,
        previousTokens: []
      },
      isSimilarSuggestionsLoading: false,
      similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
      selectedSimilarSuggestion: undefined,
      isSelectedSimilarSuggestionLoading: false,
      myVoteSuggestions: [],
      isMyVotesLoading: false,
      adminSuggestions: [],
      isAdminSuggestionsLoading: false,
      selectedMainTab: 'active',
      expandedCommentIds: [],
      loadingCommentIds: [],
      commentDrafts: {},
      commentComposerIds: [],
      submittingCommentIds: []
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._initialize();
    this._initializeStickyOffsets();
  }

  private _initializeStickyOffsets(): void {
    if (typeof window === 'undefined') {
      return;
    }

    this._scheduleStickyOffsetUpdate();
    window.addEventListener('resize', this._scheduleStickyOffsetUpdate);

    this._scrollContainer = this._getScrollContainer(this._rootRef.current);
    if (this._scrollContainer && this._scrollContainer !== window) {
      this._scrollContainer.addEventListener('scroll', this._scheduleStickyOffsetUpdate);
    } else {
      window.addEventListener('scroll', this._scheduleStickyOffsetUpdate);
    }
  }

  private _teardownStickyOffsets(): void {
    if (typeof window !== 'undefined') {
      window.removeEventListener('resize', this._scheduleStickyOffsetUpdate);
      if (this._scrollContainer && this._scrollContainer !== window) {
        this._scrollContainer.removeEventListener('scroll', this._scheduleStickyOffsetUpdate);
      } else {
        window.removeEventListener('scroll', this._scheduleStickyOffsetUpdate);
      }

      if (this._stickyUpdateFrame !== undefined) {
        window.cancelAnimationFrame(this._stickyUpdateFrame);
        this._stickyUpdateFrame = undefined;
      }
    }

    this._scrollContainer = undefined;
  }

  private _scheduleStickyOffsetUpdate = (): void => {
    if (typeof window === 'undefined') {
      return;
    }

    if (this._stickyUpdateFrame !== undefined) {
      window.cancelAnimationFrame(this._stickyUpdateFrame);
    }

    this._stickyUpdateFrame = window.requestAnimationFrame(() => {
      this._stickyUpdateFrame = undefined;
      this._updateStickyOffsets();
    });
  };

  private _updateStickyOffsets(): void {
    const root: HTMLElement | null = this._rootRef.current;
    const header: HTMLElement | null = this._headerRef.current;

    if (!root || !header || typeof window === 'undefined') {
      return;
    }

    const headerHeight: number = Math.ceil(header.getBoundingClientRect().height);
    root.style.setProperty('--samverkansportalen-header-height', `${headerHeight}px`);

    const pivotLinks: HTMLElement | null = this._getPivotLinksElement(root);
    const pivotLinksHeight: number = pivotLinks
      ? Math.ceil(pivotLinks.getBoundingClientRect().height)
      : 0;

    const rootStyles: CSSStyleDeclaration = window.getComputedStyle(root);
    const stickyOffsetBase: number = this._getNumericCssValue(
      rootStyles.getPropertyValue('--samverkansportalen-sticky-offset')
    );
    const stickyOffset: number = stickyOffsetBase + this._getScrollContainerOffset();
    const paddingTop: number = this._getNumericCssValue(rootStyles.paddingTop);
    const paddingBottom: number = this._getNumericCssValue(rootStyles.paddingBottom);
    const scrollContainerHeight: number =
      this._scrollContainer instanceof HTMLElement ? this._scrollContainer.clientHeight : window.innerHeight;
    const availableScrollHeight: number = Math.max(
      0,
      Math.floor(scrollContainerHeight - paddingTop - paddingBottom - headerHeight - pivotLinksHeight)
    );
    if (availableScrollHeight > 0) {
      root.style.setProperty(
        '--samverkansportalen-tab-scroll-height',
        `${availableScrollHeight}px`
      );
    } else {
      root.style.removeProperty('--samverkansportalen-tab-scroll-height');
    }
    const totalStickyHeight: number = headerHeight + pivotLinksHeight;
    const rootRect: DOMRect = root.getBoundingClientRect();
    const shouldStick: boolean =
      totalStickyHeight > 0 &&
      rootRect.top <= stickyOffset &&
      rootRect.bottom >= stickyOffset + totalStickyHeight;

    if (!shouldStick) {
      this._resetStickyStyles(root, header, pivotLinks);
      return;
    }

    this._applyStickyStyles(root, header, pivotLinks, {
      headerHeight,
      pivotLinksHeight
    });
  }

  private _getPivotLinksElement(root: HTMLElement): HTMLElement | null {
    if (this._pivotLinksElement && root.contains(this._pivotLinksElement)) {
      return this._pivotLinksElement;
    }

    const tabListSelector: string = '[data-samverkansportalen-tablist="true"]';
    const explicitTabList: HTMLElement | null = root.querySelector(tabListSelector) as HTMLElement | null;
    if (explicitTabList) {
      this._pivotLinksElement = explicitTabList;
      return explicitTabList;
    }

    const pivotRoot: Element | null = root.querySelector('.ms-Pivot');
    const searchRoot: Element | HTMLElement = pivotRoot ?? root;
    let pivotLinks: HTMLElement | null =
      (searchRoot.querySelector('[data-samverkansportalen-tablist="true"]') as HTMLElement | null) ??
      (searchRoot.querySelector('.fui-TabList') as HTMLElement | null) ??
      (searchRoot.querySelector('[role="tablist"]') as HTMLElement | null);

    if (!pivotLinks) {
      const tabLists: NodeListOf<HTMLElement> = document.querySelectorAll(
        '[data-samverkansportalen-tablist="true"], .fui-TabList, [role="tablist"]'
      );
      const rootClassSelector: string = `.${styles.samverkansportalen}`;
      tabLists.forEach((element) => {
        if (!pivotLinks && element.closest(rootClassSelector) === root) {
          pivotLinks = element;
        }
      });
    }

    if (!pivotLinks) {
      pivotLinks = searchRoot.querySelector('[role="tablist"]') as HTMLElement | null;
    }

    this._pivotLinksElement = pivotLinks ?? undefined;
    return pivotLinks;
  }

  private _getNumericCssValue(value: string): number {
    const parsed: number = parseFloat(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  private _getScrollContainer(element: HTMLElement | null): HTMLElement | Window {
    if (!element || typeof window === 'undefined' || typeof document === 'undefined') {
      return window;
    }

    const explicitScrollContainer: Element | null = document.querySelector(
      `.${SCROLL_CONTAINER_CLASS}`
    );
    if (
      explicitScrollContainer instanceof HTMLElement &&
      explicitScrollContainer.contains(element)
    ) {
      return explicitScrollContainer;
    }

    let current: HTMLElement | null = element.parentElement;
    while (current && current !== document.body) {
      const styles: CSSStyleDeclaration = window.getComputedStyle(current);
      if (
        this._isScrollableValue(styles.overflowY) ||
        this._isScrollableValue(styles.overflowX) ||
        this._isScrollableValue(styles.overflow)
      ) {
        return current;
      }

      current = current.parentElement;
    }

    return window;
  }

  private _isScrollableValue(value: string): boolean {
    return value === 'auto' || value === 'scroll' || value === 'overlay';
  }

  private _getScrollContainerOffset(): number {
    if (
      !this._scrollContainer ||
      this._scrollContainer === window ||
      !(this._scrollContainer instanceof HTMLElement)
    ) {
      return 0;
    }

    const rect: DOMRect = this._scrollContainer.getBoundingClientRect();
    return rect.top > 0 ? rect.top : 0;
  }

  private _applyStickyStyles(
    root: HTMLElement,
    header: HTMLElement,
    pivotLinks: HTMLElement | null,
    options: {
      headerHeight: number;
      pivotLinksHeight: number;
    }
  ): void {
    this._isStickyActive = true;
    root.style.setProperty('--samverkansportalen-sticky-header-padding', `${options.headerHeight}px`);
    root.style.setProperty('--samverkansportalen-sticky-tabs-padding', `${options.pivotLinksHeight}px`);

    this._applyFixedStyle(header, 7);

    if (pivotLinks) {
      this._applyFixedStyle(pivotLinks, 6);
    }
  }

  private _resetStickyStyles(
    root: HTMLElement,
    header: HTMLElement,
    pivotLinks: HTMLElement | null
  ): void {
    if (!this._isStickyActive) {
      return;
    }

    this._isStickyActive = false;
    root.style.setProperty('--samverkansportalen-sticky-header-padding', '0px');
    root.style.setProperty('--samverkansportalen-sticky-tabs-padding', '0px');
    this._clearFixedStyle(header);

    if (pivotLinks) {
      this._clearFixedStyle(pivotLinks);
    }
  }

  private _applyFixedStyle(element: HTMLElement, zIndex: number): void {
    element.style.position = 'sticky';
    element.style.zIndex = zIndex.toString();
    element.style.top = '';
    element.style.left = '';
    element.style.width = '';
    element.style.boxSizing = '';
  }

  private _clearFixedStyle(element: HTMLElement): void {
    element.style.position = '';
    element.style.top = '';
    element.style.left = '';
    element.style.width = '';
    element.style.zIndex = '';
    element.style.boxSizing = '';
  }

  private _deriveStatusStateFromProps(
    props: ISamverkansportalenProps
  ): { statuses: string[]; completedStatus: string; deniedStatus?: string; defaultStatus: string } {
    return this._deriveStatusState(
      props.statuses,
      props.completedStatus,
      props.defaultStatus,
      props.deniedStatus
    );
  }

  private _deriveStatusState(
    statusesInput: string[] | undefined,
    completedStatusCandidate: string | undefined,
    defaultStatusCandidate?: string,
    deniedStatusCandidate?: string
  ): { statuses: string[]; completedStatus: string; deniedStatus?: string; defaultStatus: string } {
    const statuses: string[] = this._sanitizeStatuses(statusesInput);
    const completedStatus: string = this._resolveCompletedStatus(completedStatusCandidate, statuses);
    const deniedStatus: string | undefined = this._resolveDeniedStatus(
      deniedStatusCandidate,
      statuses,
      completedStatus
    );
    const defaultStatus: string = this._resolveDefaultStatus(
      defaultStatusCandidate,
      statuses,
      completedStatus
    );
    return { statuses, completedStatus, deniedStatus, defaultStatus };
  }

  private _sanitizeStatuses(values: string[] | undefined): string[] {
    const source: string[] = Array.isArray(values) ? values : [];
    const seen: Set<string> = new Set();
    const results: string[] = [];

    source.forEach((entry) => {
      const normalized: string = typeof entry === 'string' ? entry.trim() : '';

      if (!normalized) {
        return;
      }

      const key: string = normalized.toLowerCase();

      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      results.push(normalized);
    });

    if (results.length === 0) {
      return ['Active', 'Done'];
    }

    return results;
  }

  private _resolveCompletedStatus(candidate: string | undefined, statuses: string[]): string {
    if (statuses.length === 0) {
      return 'Done';
    }

    const normalizedCandidate: string = typeof candidate === 'string' ? candidate.trim() : '';

    if (normalizedCandidate.length > 0) {
      const match: string | undefined = statuses.find((status) =>
        this._areStatusesEqual(status, normalizedCandidate)
      );

      if (match) {
        return match;
      }
    }

    return statuses[statuses.length - 1];
  }

  private _resolveDeniedStatus(
    candidate: string | undefined,
    statuses: string[],
    completedStatus: string
  ): string | undefined {
    if (statuses.length === 0) {
      return undefined;
    }

    const normalizedCandidate: string = typeof candidate === 'string' ? candidate.trim() : '';

    if (!normalizedCandidate) {
      return undefined;
    }

    const match: string | undefined = statuses.find((status) =>
      this._areStatusesEqual(status, normalizedCandidate)
    );

    if (!match) {
      return normalizedCandidate;
    }

    if (this._areStatusesEqual(match, completedStatus)) {
      return match;
    }

    return match;
  }

  private _resolveDefaultStatus(
    candidate: string | undefined,
    statuses: string[],
    completedStatus: string
  ): string {
    if (statuses.length === 0) {
      return completedStatus;
    }

    const normalizedCandidate: string = typeof candidate === 'string' ? candidate.trim() : '';

    if (normalizedCandidate.length > 0) {
      const match: string | undefined = statuses.find((status) =>
        this._areStatusesEqual(status, normalizedCandidate)
      );

      if (match) {
        return match;
      }
    }

    const firstActive: string | undefined = statuses.find(
      (status) => !this._areStatusesEqual(status, completedStatus)
    );

    return firstActive ?? completedStatus;
  }

  private _areStatusesEqual(left: string | undefined, right: string | undefined): boolean {
    const normalizedLeft: string = typeof left === 'string' ? left.trim().toLowerCase() : '';
    const normalizedRight: string = typeof right === 'string' ? right.trim().toLowerCase() : '';
    return normalizedLeft === normalizedRight;
  }

  private _isCompletedStatusValue(
    status: string | undefined,
    completedStatus: string,
    deniedStatus?: string
  ): boolean {
    if (this._areStatusesEqual(status, completedStatus)) {
      return true;
    }

    return this._isDeniedStatusValue(status, deniedStatus);
  }

  private _isDeniedStatusValue(status: string | undefined, deniedStatus: string | undefined): boolean {
    if (!deniedStatus) {
      return false;
    }

    return this._areStatusesEqual(status, deniedStatus);
  }

  private _filterDeniedSuggestions(items: ISuggestionItem[]): ISuggestionItem[] {
    if (
      !this.state.deniedStatus ||
      this._areStatusesEqual(this.state.deniedStatus, this.state.completedStatus)
    ) {
      return items;
    }

    return items.filter(
      (item) => !this._isDeniedStatusValue(item.status, this.state.deniedStatus)
    );
  }

  private _normalizeActiveStatusValue(
    status: string | undefined,
    statuses: string[],
    completedStatus: string,
    deniedStatus?: string
  ): string | undefined {
    if (!status) {
      return undefined;
    }

    const match: string | undefined = statuses.find((entry) => this._areStatusesEqual(entry, status));

    if (!match) {
      return undefined;
    }

    if (this._isCompletedStatusValue(match, completedStatus, deniedStatus)) {
      return undefined;
    }

    return match;
  }

  private _areStatusCollectionsEqual(left: string[] | undefined, right: string[] | undefined): boolean {
    const leftNormalized: string[] = this._sanitizeStatuses(left);
    const rightNormalized: string[] = this._sanitizeStatuses(right);

    if (leftNormalized.length !== rightNormalized.length) {
      return false;
    }

    return leftNormalized.every((value, index) =>
      this._areStatusesEqual(value, rightNormalized[index])
    );
  }

  private _applyStatusConfiguration(
    statusesOverride?: string[],
    completedStatusOverride?: string,
    options: { reloadSuggestions?: boolean; defaultStatusOverride?: string; deniedStatusOverride?: string } = {}
  ): void {
    const { reloadSuggestions = true, defaultStatusOverride } = options;
    const { statuses, completedStatus, deniedStatus, defaultStatus } = this._deriveStatusState(
      statusesOverride ?? this.props.statuses,
      completedStatusOverride ?? this.props.completedStatus,
      defaultStatusOverride ?? this.props.defaultStatus,
      options.deniedStatusOverride ?? this.props.deniedStatus
    );
    const shouldReload: boolean = reloadSuggestions;

    this._updateState(
      (prevState) => {
        const activeStatus: string | undefined = this._normalizeActiveStatusValue(
          prevState.activeFilter.status,
          statuses,
          completedStatus,
          deniedStatus
        );
        const adminStatus: string | undefined = this._normalizeActiveStatusValue(
          prevState.adminFilter.status,
          statuses,
          completedStatus,
          deniedStatus
        );
        const includeDenied: boolean =
          prevState.completedFilter.includeDenied === true &&
          !!deniedStatus &&
          !this._areStatusesEqual(deniedStatus, completedStatus);

        return {
          statuses,
          completedStatus,
          deniedStatus,
          defaultStatus,
          activeFilter: { ...prevState.activeFilter, status: activeStatus },
          completedFilter: {
            ...prevState.completedFilter,
            status: completedStatus,
            includeDenied
          },
          adminFilter: { ...prevState.adminFilter, status: adminStatus }
        } as Partial<ISamverkansportalenState>;
      },
      () => {
        if (shouldReload) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._loadSuggestions();
        }
      }
    );
  }

  private _getActiveStatuses(): string[] {
    return this.state.statuses.filter(
      (status) =>
        !this._isCompletedStatusValue(status, this.state.completedStatus, this.state.deniedStatus)
    );
  }

  private _getCompletedStatuses(filter: IFilterState): string[] {
    const statuses: string[] = [this.state.completedStatus];

    if (
      filter.includeDenied &&
      this.state.deniedStatus &&
      !this._areStatusesEqual(this.state.deniedStatus, this.state.completedStatus)
    ) {
      statuses.push(this.state.deniedStatus);
    }

    return statuses;
  }

  private _normalizeAdminFilterStatus(status: string | undefined): string | undefined {
    return this._normalizeActiveStatusValue(
      status,
      this.state.statuses,
      this.state.completedStatus,
      this.state.deniedStatus
    );
  }

  private _normalizeStatusValue(status: string | undefined, statuses: string[]): string | undefined {
    if (!status) {
      return undefined;
    }

    return statuses.find((entry) => this._areStatusesEqual(entry, status));
  }

  private _isStatusInCollection(status: string | undefined, collection: string[]): boolean {
    if (!status) {
      return false;
    }

    return collection.some((entry) => this._areStatusesEqual(entry, status));
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this._teardownStickyOffsets();
    this._debouncedSimilarSuggestionsSearch.cancel();
    this._debouncedActiveFilterSearch.cancel();
    this._debouncedCompletedFilterSearch.cancel();
  }

  public componentDidUpdate(prevProps: ISamverkansportalenProps): void {
    const listChanged: boolean = this._normalizeListTitle(prevProps.listTitle) !== this._listTitle;
    const voteListChanged: boolean =
      this._normalizeVoteListTitle(prevProps.voteListTitle, prevProps.listTitle) !== this._voteListTitle;
    const commentListChanged: boolean =
      this._normalizeCommentListTitle(prevProps.commentListTitle, prevProps.listTitle) !== this._commentListTitle;
    const subcategoryListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.subcategoryListTitle) !== this._subcategoryListTitle;
    const categoryListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.categoryListTitle) !== this._categoryListTitle;
    const statusListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.statusListTitle) !== this._statusListTitle;
    const statusesChanged: boolean =
      !this._areStatusCollectionsEqual(prevProps.statuses, this.props.statuses) ||
      !this._areStatusesEqual(prevProps.completedStatus, this.props.completedStatus);
    const deniedStatusChanged: boolean = !this._areStatusesEqual(
      prevProps.deniedStatus,
      this.props.deniedStatus
    );
    const defaultStatusChanged: boolean = !this._areStatusesEqual(
      prevProps.defaultStatus,
      this.props.defaultStatus
    );
    const totalVotesChanged: boolean = prevProps.totalVotesPerUser !== this.props.totalVotesPerUser;

    if (statusesChanged || deniedStatusChanged || defaultStatusChanged) {
      if (this._statusListTitle) {
        this._applyStatusConfiguration(this.state.statuses, this.props.completedStatus, {
          defaultStatusOverride: this.props.defaultStatus,
          deniedStatusOverride: this.props.deniedStatus
        });
      } else {
        this._applyStatusConfiguration(undefined, undefined, {
          defaultStatusOverride: this.props.defaultStatus,
          deniedStatusOverride: this.props.deniedStatus
        });
      }
    }

    if (
      listChanged ||
      voteListChanged ||
      commentListChanged ||
      subcategoryListChanged ||
      categoryListChanged ||
      statusListChanged
    ) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._initialize();
    }

    if (totalVotesChanged) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._loadAvailableVotes();
    }

    this._scheduleStickyOffsetUpdate();
  }

  public render(): React.ReactElement<ISamverkansportalenProps> {
    const {
      activeSuggestions,
      completedSuggestions,
      similarSuggestions,
      isLoading,
      isActiveSuggestionsLoading,
      isCompletedSuggestionsLoading,
      isSimilarSuggestionsLoading,
      newTitle,
      newDescription,
      newCategory,
      newSubcategoryKey,
      subcategories,
      categories,
      activeFilter,
      completedFilter,
      similarSuggestionsQuery,
      error,
      success,
      myVoteSuggestions,
      isMyVotesLoading,
      adminSuggestions,
      isAdminSuggestionsLoading,
      adminFilter,
      selectedMainTab
    } = this.state;

    const subcategoryOptions: IDropdownOption[] = this._getSubcategoryOptions(newCategory, subcategories);
    const categoryOptions: IDropdownOption[] = this._getCategoryOptions(categories);
    const filterCategoryOptions: IDropdownOption[] = this._getFilterCategoryOptions(categories);
    const activeFilterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      activeFilter.category,
      subcategories
    );
    const completedFilterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      completedFilter.category,
      subcategories
    );
    const adminFilterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      adminFilter.category,
      subcategories
    );
    const adminFilterStatusOptions: IDropdownOption[] = this._getFilterStatusOptions();

    const isFilterCategoryLimited: boolean = filterCategoryOptions.length <= 1;
    const isActiveFilterSubcategoryLimited: boolean = activeFilterSubcategoryOptions.length <= 1;
    const isCompletedFilterSubcategoryLimited: boolean = completedFilterSubcategoryOptions.length <= 1;
    const isAdminFilterSubcategoryLimited: boolean = adminFilterSubcategoryOptions.length <= 1;
    const isAdminFilterStatusLimited: boolean = adminFilterStatusOptions.length <= 1;
    const activeFilterSubcategoryPlaceholder: string = isActiveFilterSubcategoryLimited
      ? strings.NoSubcategoriesAvailablePlaceholder
      : strings.SelectSubcategoryPlaceholder;
    const completedFilterSubcategoryPlaceholder: string = isCompletedFilterSubcategoryLimited
      ? strings.NoSubcategoriesAvailablePlaceholder
      : strings.SelectSubcategoryPlaceholder;
    const adminFilterSubcategoryPlaceholder: string = isAdminFilterSubcategoryLimited
      ? strings.NoSubcategoriesAvailablePlaceholder
      : strings.SelectSubcategoryPlaceholder;

    const hasActiveFilters: boolean = this._hasSearchFilters(activeFilter);
    const hasCompletedFilters: boolean = this._hasSearchFilters(completedFilter);
    const hasAdminFiltersApplied: boolean = this._hasAdminFilters(adminFilter);
    const showDeniedFilter: boolean =
      !!this.state.deniedStatus &&
      !this._areStatusesEqual(this.state.deniedStatus, this.state.completedStatus);

    const activeSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      activeSuggestions.items,
      false
    );
    const completedSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      completedSuggestions.items,
      true
    );
    const similarSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      similarSuggestions.items,
      true,
      { allowVoting: true }
    );
    const myVoteSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      myVoteSuggestions,
      false,
      { allowVoting: true }
    );
    const adminSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      adminSuggestions,
      false,
      { allowVoting: true }
    );
    const voteSummaryOptions: IDropdownOption[] = this._getVoteSummaryOptions(categories);
    const formatTabLabel = (label: string, total?: number): string =>
      typeof total === 'number' ? `${label} (${total})` : label;
    const addTabLabel: string = strings.AddSuggestionTabLabel;
    const activeTotalCount: number | undefined = 
      this.state.activeSuggestionsTotal ?? activeSuggestions.totalCount;
    const completedTotalCount: number | undefined =
      this.state.completedSuggestionsTotal ?? completedSuggestions.totalCount;
    const activeTabLabel: string = formatTabLabel(
      strings.ActiveSuggestionsTabLabel,
      activeTotalCount
    );
    const completedTabLabel: string = formatTabLabel(
      strings.CompletedSuggestionsTabLabel,
      completedTotalCount
    );
    const activeTotalPages: number | undefined = this._getTotalPages(
      activeTotalCount,
      this.state.activePageSize
    );
    const completedTotalPages: number | undefined = this._getTotalPages(
      completedTotalCount,
      this.state.completedPageSize
    );
    const tabItems: Array<{ key: MainTabKey; label: string }> = [
      { key: 'add', label: addTabLabel },
      { key: 'active', label: activeTabLabel },
      { key: 'completed', label: completedTabLabel },
      { key: 'admin', label: strings.AdminTopSuggestionsTabLabel },
      { key: 'myVotes', label: strings.MyVotesTabLabel }
    ];

    return (
      <section
        ref={this._rootRef}
        className={`${styles.samverkansportalen} ${this.props.hasTeamsContext ? styles.teams : ''}`}
      >
        <header ref={this._headerRef} className={styles.header}>
          <div>
            <h2 className={styles.title}>{this.props.headerTitle}</h2>
            <p className={styles.subtitle}>{this.props.headerSubtitle}</p>
          </div>
          <div className={styles.voteSummary} aria-live="polite">
            <Dropdown
              className={styles.voteSummaryDropdown}
              label={strings.VotesRemainingLabel}
              options={voteSummaryOptions}
              defaultSelectedKey={voteSummaryOptions[0]?.key}
              disabled={voteSummaryOptions.length === 0}
            />
          </div>
        </header>

        {error && (
          <MessageBar
            className={styles.messageBar}
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={this._dismissError}
          >
            {error}
          </MessageBar>
        )}

        {success && (
          <MessageBar
            className={styles.messageBar}
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={this._dismissSuccess}
          >
            {success}
          </MessageBar>
        )}


        <TabHeader items={tabItems} selectedKey={selectedMainTab} onSelect={this._onTabSelect} />

        <TabPanel tabKey="add" selectedKey={selectedMainTab}>
          <div className="samverkansportalen-tab-scroll">
            <div className={styles.pivotContent}>
              <div className={styles.addSuggestion}>
                <div className={styles.sectionHeader}>
                  <h3 id={this._sectionIds.add.title} className={styles.sectionTitle}>
                    {strings.AddSuggestionSectionTitle}
                  </h3>
                </div>
                <div
                  id={this._sectionIds.add.content}
                  role="region"
                  aria-labelledby={this._sectionIds.add.title}
                  className={styles.sectionContent}
                >
                  <div className={styles.addForm}>
                    <TextField
                      label={strings.AddSuggestionTitleLabel}
                      required
                      value={newTitle}
                      onChange={this._onTitleChange}
                      disabled={isLoading}
                    />
                    <RichTextEditor
                      label={strings.AddSuggestionDetailsLabel}
                      value={newDescription}
                      onChange={this._onDescriptionEditorChange}
                      placeholder={strings.RichTextEditorPlaceholder}
                      disabled={isLoading}
                    />
                    <Dropdown
                      label={strings.CategoryLabel}
                      options={categoryOptions}
                      selectedKey={newCategory}
                      onChange={this._onCategoryChange}
                      disabled={isLoading || categoryOptions.length === 0}
                    />
                    <Dropdown
                      label={strings.SubcategoryLabel}
                      options={subcategoryOptions}
                      selectedKey={newSubcategoryKey}
                      onChange={this._onSubcategoryChange}
                      disabled={isLoading || subcategoryOptions.length === 0}
                      placeholder={
                        subcategoryOptions.length === 0
                          ? strings.NoSubcategoriesAvailablePlaceholder
                          : strings.SelectSubcategoryPlaceholder
                      }
                    />
                    <PrimaryButton
                      text={strings.SubmitSuggestionButtonText}
                      onClick={this._addSuggestion}
                      disabled={isLoading || newTitle.trim().length === 0}
                    />
                    <SimilarSuggestions
                      viewModels={similarSuggestionViewModels}
                      isLoading={isSimilarSuggestionsLoading}
                      query={similarSuggestionsQuery}
                      page={similarSuggestions.page}
                      hasPrevious={similarSuggestions.previousTokens.length > 0}
                      hasNext={!!similarSuggestions.nextToken}
                      onPrevious={this._goToPreviousSimilarPage}
                      onNext={this._goToNextSimilarPage}
                      onToggleVote={(item) => this._toggleVote(item)}
                      onChangeStatus={(item, status) => this._updateSuggestionStatus(item, status)}
                      onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                      onSubmitComment={(item) => this._submitCommentForSuggestion(item)}
                      onCommentDraftChange={(item, value) => this._handleCommentDraftChange(item, value)}
                      onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                      onToggleComments={(id) => this._toggleCommentsSection(id)}
                      onToggleCommentComposer={(id) => this._toggleCommentComposer(id)}
                      formatDateTime={(value) => this._formatDateTime(value)}
                      isProcessing={isLoading}
                      statuses={this.state.statuses}
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </TabPanel>

        <TabPanel tabKey="active" selectedKey={selectedMainTab}>
          <div className="samverkansportalen-tab-scroll">
            <div className={styles.pivotContent}>
              <SuggestionSection
                title={strings.ActiveSuggestionsSectionTitle}
                titleId={this._sectionIds.active.title}
                contentId={this._sectionIds.active.content}
                isLoading={isLoading}
                isSectionLoading={isActiveSuggestionsLoading}
                searchValue={activeFilter.searchQuery}
                onSearchChange={this._onActiveSearchChange}
                categoryOptions={filterCategoryOptions}
                selectedCategoryKey={activeFilter.category ?? ALL_CATEGORY_FILTER_KEY}
                onCategoryChange={this._onActiveFilterCategoryChange}
                disableCategoryDropdown={isFilterCategoryLimited}
                subcategoryOptions={activeFilterSubcategoryOptions}
                selectedSubcategoryKey={activeFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                onSubcategoryChange={this._onActiveFilterSubcategoryChange}
                disableSubcategoryDropdown={isActiveFilterSubcategoryLimited}
                subcategoryPlaceholder={activeFilterSubcategoryPlaceholder}
                onClearFilters={this._clearActiveFilters}
                isClearFiltersDisabled={!hasActiveFilters}
                pageSizeOptions={SUGGESTION_PAGE_SIZE_OPTIONS}
                selectedPageSize={this.state.activePageSize}
                onPageSizeChange={this._onActivePageSizeChange}
                viewModels={activeSuggestionViewModels}
                useTableLayout={this.props.useTableLayout === true}
                showMetadataInIdColumn={this.props.showMetadataInIdColumn === true}
                totalPages={activeTotalPages}
                onToggleVote={(item) => this._toggleVote(item)}
                onChangeStatus={(item, status) => this._updateSuggestionStatus(item, status)}
                onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                onSubmitComment={(item) => this._submitCommentForSuggestion(item)}
                onCommentDraftChange={(item, value) => this._handleCommentDraftChange(item, value)}
                onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                onToggleComments={(id) => this._toggleCommentsSection(id)}
                onToggleCommentComposer={(id) => this._toggleCommentComposer(id)}
                formatDateTime={(value) => this._formatDateTime(value)}
                statuses={this.state.statuses}
                page={activeSuggestions.page}
                hasPrevious={activeSuggestions.page > 1}
                hasNext={!!activeSuggestions.nextToken}
                onPrevious={this._goToPreviousActivePage}
                onNext={this._goToNextActivePage}
              />
            </div>
          </div>
        </TabPanel>

        <TabPanel tabKey="completed" selectedKey={selectedMainTab}>
          <div className="samverkansportalen-tab-scroll">
            <div className={styles.pivotContent}>
              <SuggestionSection
                title={strings.CompletedSuggestionsSectionTitle}
                titleId={this._sectionIds.completed.title}
                contentId={this._sectionIds.completed.content}
                isLoading={isLoading}
                isSectionLoading={isCompletedSuggestionsLoading}
                searchValue={completedFilter.searchQuery}
                onSearchChange={this._onCompletedSearchChange}
                categoryOptions={filterCategoryOptions}
                selectedCategoryKey={completedFilter.category ?? ALL_CATEGORY_FILTER_KEY}
                onCategoryChange={this._onCompletedFilterCategoryChange}
                disableCategoryDropdown={isFilterCategoryLimited}
                subcategoryOptions={completedFilterSubcategoryOptions}
                selectedSubcategoryKey={completedFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                onSubcategoryChange={this._onCompletedFilterSubcategoryChange}
                disableSubcategoryDropdown={isCompletedFilterSubcategoryLimited}
                subcategoryPlaceholder={completedFilterSubcategoryPlaceholder}
                showDeniedFilter={showDeniedFilter}
                isDeniedFilterOn={completedFilter.includeDenied === true}
                onDeniedFilterChange={this._onCompletedDeniedFilterChange}
                onClearFilters={this._clearCompletedFilters}
                isClearFiltersDisabled={!hasCompletedFilters}
                pageSizeOptions={SUGGESTION_PAGE_SIZE_OPTIONS}
                selectedPageSize={this.state.completedPageSize}
                onPageSizeChange={this._onCompletedPageSizeChange}
                viewModels={completedSuggestionViewModels}
                useTableLayout={this.props.useTableLayout === true}
                showMetadataInIdColumn={this.props.showMetadataInIdColumn === true}
                totalPages={completedTotalPages}
                onToggleVote={(item) => this._toggleVote(item)}
                onChangeStatus={(item, status) => this._updateSuggestionStatus(item, status)}
                onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                onSubmitComment={(item) => this._submitCommentForSuggestion(item)}
                onCommentDraftChange={(item, value) => this._handleCommentDraftChange(item, value)}
                onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                onToggleComments={(id) => this._toggleCommentsSection(id)}
                onToggleCommentComposer={(id) => this._toggleCommentComposer(id)}
                formatDateTime={(value) => this._formatDateTime(value)}
                statuses={this.state.statuses}
                page={completedSuggestions.page}
                hasPrevious={completedSuggestions.page > 1}
                hasNext={!!completedSuggestions.nextToken}
                onPrevious={this._goToPreviousCompletedPage}
                onNext={this._goToNextCompletedPage}
              />
            </div>
          </div>
        </TabPanel>

        <TabPanel tabKey="admin" selectedKey={selectedMainTab}>
          <div className="samverkansportalen-tab-scroll">
            <div className={styles.pivotContent}>
              <div className={styles.filters}>
                <div className={styles.filterControls}>
                  <Dropdown
                    label={strings.CategoryLabel}
                    options={filterCategoryOptions}
                    selectedKey={adminFilter.category ?? ALL_CATEGORY_FILTER_KEY}
                    onChange={this._onAdminFilterCategoryChange}
                    disabled={isFilterCategoryLimited}
                    className={styles.filterDropdown}
                  />
                  <Dropdown
                    label={strings.SubcategoryLabel}
                    options={adminFilterSubcategoryOptions}
                    selectedKey={adminFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                    onChange={this._onAdminFilterSubcategoryChange}
                    disabled={isAdminFilterSubcategoryLimited}
                    placeholder={adminFilterSubcategoryPlaceholder}
                    className={styles.filterDropdown}
                  />
                  <Dropdown
                    label={strings.StatusLabel}
                    options={adminFilterStatusOptions}
                    selectedKey={adminFilter.status ?? ALL_STATUS_FILTER_KEY}
                    onChange={this._onAdminFilterStatusChange}
                    disabled={isAdminFilterStatusLimited}
                    className={styles.filterDropdown}
                  />
                  <DefaultButton
                    text={strings.ClearFiltersButtonText}
                    className={styles.filterButton}
                    onClick={this._clearAdminFilters}
                    disabled={isLoading || isAdminSuggestionsLoading || !hasAdminFiltersApplied}
                  />
                </div>
              </div>

              {isAdminSuggestionsLoading ? (
                <Spinner label={strings.LoadingSuggestionsLabel} size={SpinnerSize.large} />
              ) : adminSuggestions.length === 0 ? (
                <p className={styles.emptyState}>{strings.NoSuggestionsLabel}</p>
              ) : (
                <SuggestionList
                  viewModels={adminSuggestionViewModels}
                  useTableLayout={this.props.useTableLayout === true}
                  showMetadataInIdColumn={this.props.showMetadataInIdColumn === true}
                  isLoading={isLoading || isAdminSuggestionsLoading}
                  onToggleVote={(item) => this._toggleVote(item)}
                  onChangeStatus={(item, status) => this._updateSuggestionStatus(item, status)}
                  onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                  onSubmitComment={(item) => this._submitCommentForSuggestion(item)}
                  onCommentDraftChange={(item, value) => this._handleCommentDraftChange(item, value)}
                  onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                  onToggleComments={(id) => this._toggleCommentsSection(id)}
                  onToggleCommentComposer={(id) => this._toggleCommentComposer(id)}
                  formatDateTime={(value) => this._formatDateTime(value)}
                  statuses={this.state.statuses}
                />
              )}
            </div>
          </div>
        </TabPanel>

        <TabPanel tabKey="myVotes" selectedKey={selectedMainTab}>
          <div className="samverkansportalen-tab-scroll">
            <div className={styles.pivotContent}>
              {isMyVotesLoading ? (
                <Spinner label={strings.LoadingSuggestionsLabel} size={SpinnerSize.large} />
              ) : myVoteSuggestions.length === 0 ? (
                <p className={styles.emptyState}>{strings.NoVotedSuggestionsLabel}</p>
              ) : (
                <SuggestionList
                  viewModels={myVoteSuggestionViewModels}
                  useTableLayout={this.props.useTableLayout === true}
                  showMetadataInIdColumn={this.props.showMetadataInIdColumn === true}
                  isLoading={isLoading || isMyVotesLoading}
                  onToggleVote={(item) => this._toggleVote(item)}
                  onChangeStatus={(item, status) => this._updateSuggestionStatus(item, status)}
                  onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                  onSubmitComment={(item) => this._submitCommentForSuggestion(item)}
                  onCommentDraftChange={(item, value) => this._handleCommentDraftChange(item, value)}
                  onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                  onToggleComments={(id) => this._toggleCommentsSection(id)}
                  onToggleCommentComposer={(id) => this._toggleCommentComposer(id)}
                  formatDateTime={(value) => this._formatDateTime(value)}
                  statuses={this.state.statuses}
                />
              )}
            </div>
          </div>
        </TabPanel>
      </section>
    );
  }

  private _createSuggestionViewModels(
    items: ISuggestionItem[],
    readOnly: boolean,
    options: { allowVoting?: boolean } = {}
  ): ISuggestionViewModel[] {
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);
    const allowVoting: boolean = options.allowVoting === true;

    return items.map((item) => {
      const isCompleted: boolean = this._isCompletedStatusValue(
        item.status,
        this.state.completedStatus,
        this.state.deniedStatus
      );
      const remainingVotesForCategory: number = this._getRemainingVotesForCategory(item.category);
      const interaction: ISuggestionInteractionState = this._getInteractionState(
        item,
        readOnly,
        normalizedUser,
        remainingVotesForCategory,
        allowVoting
      );
      const isExpanded: boolean = this._isCommentSectionExpanded(item.id);
      const isLoadingComments: boolean = this.state.loadingCommentIds.indexOf(item.id) !== -1;
      const hasLoadedComments: boolean = item.areCommentsLoaded;
      const resolvedCommentCount: number = hasLoadedComments ? item.comments.length : item.commentCount;
      const renderedComments: ISuggestionComment[] = hasLoadedComments ? item.comments : [];
      const regionId: string = `${this._commentSectionPrefix}-${item.id}`;
      const toggleId: string = `${regionId}-toggle`;
      const isComposerVisible: boolean = this._isCommentComposerVisible(item.id);
      const draftText: string = this._getCommentDraft(item.id);
      const isSubmittingComment: boolean = this._isCommentSubmitting(item.id);

      return {
        item,
        interaction,
        comment: {
          isExpanded,
          isLoading: isLoadingComments,
          hasLoaded: hasLoadedComments,
          resolvedCount: resolvedCommentCount,
          comments: renderedComments,
          canAddComment: interaction.canAddComment,
          canDeleteComments: this.props.isCurrentUserAdmin && !isCompleted,
          regionId,
          toggleId,
          isComposerVisible,
          draftText,
          isSubmitting: isSubmittingComment
        }
      };
    });
  }

  private _goToPreviousActivePage = async (): Promise<void> => {
    const { activeSuggestions, activeFilter } = this.state;

    if (activeSuggestions.page <= 1) {
      return;
    }

    const tokens: (string | undefined)[] = [...activeSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchActiveSuggestions({
        page: Math.max(activeSuggestions.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter: activeFilter
      });
    } catch (error) {
      this._handleError(strings.ActiveSuggestionsPreviousPageErrorMessage, error);
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  };

  private _goToNextActivePage = async (): Promise<void> => {
    const { activeSuggestions, activeFilter } = this.state;

    if (!activeSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...activeSuggestions.previousTokens,
      ...(activeSuggestions.currentToken ? [activeSuggestions.currentToken] : [])
    ];

    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchActiveSuggestions({
        page: activeSuggestions.page + 1,
        previousTokens: tokens,
        skipToken: activeSuggestions.nextToken,
        filter: activeFilter
      });
    } catch (error) {
      this._handleError(strings.ActiveSuggestionsNextPageErrorMessage, error);
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  };

  private _goToPreviousCompletedPage = async (): Promise<void> => {
    const { completedSuggestions, completedFilter } = this.state;

    if (completedSuggestions.page <= 1) {
      return;
    }

    const tokens: (string | undefined)[] = [...completedSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchCompletedSuggestions({
        page: Math.max(completedSuggestions.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter: completedFilter
      });
    } catch (error) {
      this._handleError(strings.CompletedSuggestionsPreviousPageErrorMessage, error);
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  };

  private _goToNextCompletedPage = async (): Promise<void> => {
    const { completedSuggestions, completedFilter } = this.state;

    if (!completedSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...completedSuggestions.previousTokens,
      ...(completedSuggestions.currentToken ? [completedSuggestions.currentToken] : [])
    ];

    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchCompletedSuggestions({
        page: completedSuggestions.page + 1,
        previousTokens: tokens,
        skipToken: completedSuggestions.nextToken,
        filter: completedFilter
      });
    } catch (error) {
      this._handleError(strings.CompletedSuggestionsNextPageErrorMessage, error);
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  };

  private _goToPreviousSimilarPage = async (): Promise<void> => {
    const { similarSuggestions, similarSuggestionsQuery } = this.state;

    if (similarSuggestions.previousTokens.length === 0) {
      return;
    }

    const tokens: (string | undefined)[] = [...similarSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    await this._fetchSimilarSuggestions({
      page: Math.max(similarSuggestions.page - 1, 1),
      previousTokens: tokens,
      skipToken: previousToken,
      query: { ...similarSuggestionsQuery }
    });
  };

  private _goToNextSimilarPage = async (): Promise<void> => {
    const { similarSuggestions, similarSuggestionsQuery } = this.state;

    if (!similarSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...similarSuggestions.previousTokens,
      similarSuggestions.currentToken
    ];

    await this._fetchSimilarSuggestions({
      page: similarSuggestions.page + 1,
      previousTokens: tokens,
      skipToken: similarSuggestions.nextToken,
      query: { ...similarSuggestionsQuery }
    });
  };

  private _getInteractionState(
    item: ISuggestionItem,
    readOnly: boolean,
    normalizedUser: string | undefined,
    remainingVotesForCategory: number,
    allowVoting: boolean
  ): {
    hasVoted: boolean;
    disableVote: boolean;
    canAddComment: boolean;
    canAdvanceSuggestionStatus: boolean;
    canDeleteSuggestion: boolean;
    isVotingAllowed: boolean;
  } {
    const hasVoted: boolean = !!normalizedUser && item.voters.indexOf(normalizedUser) !== -1;
    const isCompleted: boolean = this._isCompletedStatusValue(
      item.status,
      this.state.completedStatus,
      this.state.deniedStatus
    );
    const isVotingAllowed: boolean = !isCompleted && (allowVoting || !readOnly);
    const disableVote: boolean =
      this.state.isLoading ||
      !isVotingAllowed ||
      (!hasVoted && !this.state.isUnlimitedVotes && remainingVotesForCategory <= 0);
    const canAdvanceSuggestionStatus: boolean = this.props.isCurrentUserAdmin && !readOnly && !isCompleted;
    const canDeleteSuggestion: boolean = this._canCurrentUserDeleteSuggestion(item);
    const canAddComment: boolean = !readOnly && !isCompleted;

    return {
      hasVoted,
      disableVote,
      canAddComment,
      canAdvanceSuggestionStatus,
      canDeleteSuggestion,
      isVotingAllowed
    };
  }

  private _formatDateTime(value: string): string {
    try {
      const parsed: Date = new Date(value);

      if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleString();
      }
    } catch (error) {
      console.warn('Failed to parse completion date.', error);
    }

    return value;
  }

  private _getSubcategoryOptions(
    category: SuggestionCategory,
    definitions: ISubcategoryDefinition[]
  ): IDropdownOption[] {
    return this._getSubcategoriesForCategory(category, definitions).map((definition) => ({
      key: definition.key,
      text: definition.title
    }));
  }

  private _getFilterSubcategoryOptions(
    category: SuggestionCategory | undefined,
    definitions: ISubcategoryDefinition[]
  ): IDropdownOption[] {
    const options: IDropdownOption[] = this._getSubcategoriesForCategory(category, definitions).map(
      (definition) => ({
        key: definition.title,
        text: definition.title
      })
    );

    return [{ key: ALL_SUBCATEGORY_FILTER_KEY, text: strings.AllSubcategoriesOptionLabel }, ...options];
  }

  private _getCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return categories.map((category) => ({ key: category, text: category }));
  }

  private _getFilterCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return [{ key: ALL_CATEGORY_FILTER_KEY, text: strings.AllCategoriesOptionLabel }, ...this._getCategoryOptions(categories)];
  }

  private _getFilterStatusOptions(): IDropdownOption[] {
    const options: IDropdownOption[] = this._getActiveStatuses().map((status) => ({
      key: status,
      text: status
    }));

    return [{ key: ALL_STATUS_FILTER_KEY, text: strings.AllStatusesOptionLabel }, ...options];
  }

  private _getSubcategoriesForCategory(
    category: SuggestionCategory | undefined,
    definitions: ISubcategoryDefinition[] = this.state.subcategories
  ): ISubcategoryDefinition[] {
    return definitions.filter((definition) => !definition.category || !category || definition.category === category);
  }

  private _normalizeFilterSubcategory(
    category: SuggestionCategory | undefined,
    preferredSubcategory: string | undefined,
    definitions: ISubcategoryDefinition[]
  ): string | undefined {
    if (!preferredSubcategory) {
      return undefined;
    }

    const availableTitles: string[] = this._getSubcategoriesForCategory(category, definitions).map((definition) =>
      definition.title.trim()
    );

    return availableTitles.indexOf(preferredSubcategory) !== -1 ? preferredSubcategory : undefined;
  }

  private _getValidSubcategoryKeyForCategory(
    category: SuggestionCategory,
    preferredKey: string | undefined,
    definitions: ISubcategoryDefinition[] = this.state.subcategories
  ): string | undefined {
    const options: ISubcategoryDefinition[] = this._getSubcategoriesForCategory(category, definitions);

    if (preferredKey && options.some((option) => option.key === preferredKey)) {
      return preferredKey;
    }

    return options.length > 0 ? options[0].key : undefined;
  }

  private _getSelectedSubcategoryDefinition(): ISubcategoryDefinition | undefined {
    const { newSubcategoryKey, subcategories } = this.state;

    if (!newSubcategoryKey) {
      return undefined;
    }

    return subcategories.find((definition) => definition.key === newSubcategoryKey);
  }

  private async _initialize(): Promise<void> {
    this._currentListId = undefined;
    this._currentVotesListId = undefined;
    this._currentCommentsListId = undefined;
    this._currentSubcategoryListId = undefined;
    this._currentCategoryListId = undefined;
    this._currentStatusListId = undefined;
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      await this._ensureLists();
      await this._ensureCategoryList();
      await this._ensureSubcategoryList();
      await this._ensureStatusList();
      await this._loadSuggestions();
    } catch (error) {
      const message: string =
        error instanceof Error && error.message.includes('category list')
          ? strings.ConfiguredCategoryLoadErrorMessage
          : error instanceof Error && error.message.includes('subcategory list')
          ? strings.ConfiguredSubcategoryLoadErrorMessage
          : error instanceof Error && error.message.includes('status list')
          ? strings.StatusListLoadErrorMessage
          : strings.SuggestionsListLoadErrorMessage;
      this._handleError(message, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _ensureLists(): Promise<void> {
    const listTitle: string = this._listTitle;
    const voteListTitle: string = this._voteListTitle;
    const commentListTitle: string = this._commentListTitle;
    const result = await this.props.graphService.ensureList(listTitle);
    this._currentListId = result.id;
    this._flushPendingSimilarSuggestionsSearch();

    const votesResult = await this.props.graphService.ensureVoteList(voteListTitle);
    this._currentVotesListId = votesResult.id;

    const commentsResult = await this.props.graphService.ensureCommentList(commentListTitle);
    this._currentCommentsListId = commentsResult.id;
  }

  private async _ensureCategoryList(): Promise<void> {
    this._currentCategoryListId = undefined;

    const listTitle: string | undefined = this._categoryListTitle;

    if (!listTitle) {
      this._applyCategories(FALLBACK_CATEGORIES);
      return;
    }

    const listInfo = await this.props.graphService.getListByTitle(listTitle);

    if (!listInfo) {
      throw new Error(`Failed to load the category list "${listTitle}".`);
    }

    this._currentCategoryListId = listInfo.id;
    await this._loadCategories();
  }

  private async _ensureSubcategoryList(): Promise<void> {
    this._currentSubcategoryListId = undefined;
    this._updateState({ subcategories: [], newSubcategoryKey: undefined });

    const listTitle: string | undefined = this._subcategoryListTitle;

    if (!listTitle) {
      return;
    }

    const listInfo = await this.props.graphService.getListByTitle(listTitle);

    if (!listInfo) {
      throw new Error(`Failed to load the subcategory list "${listTitle}".`);
    }

    this._currentSubcategoryListId = listInfo.id;
    await this._loadSubcategories();
  }

  private async _ensureStatusList(): Promise<void> {
    this._currentStatusListId = undefined;

    const listTitle: string | undefined = this._statusListTitle;

    if (!listTitle) {
      this._applyStatusConfiguration(undefined, undefined, { reloadSuggestions: false });
      return;
    }

    const listInfo = await this.props.graphService.getListByTitle(listTitle);

    if (!listInfo) {
      throw new Error(`Failed to load the status list "${listTitle}".`);
    }

    this._currentStatusListId = listInfo.id;
    await this._loadStatusesFromList();
  }

  private async _loadSuggestions(): Promise<void> {
    this._updateState({
      isActiveSuggestionsLoading: true,
      isCompletedSuggestionsLoading: true
    });

    try {
      await Promise.all([
        this._fetchActiveSuggestions({
          page: 1,
          previousTokens: [],
          skipToken: undefined,
          filter: this.state.activeFilter
        }),
        this._fetchCompletedSuggestions({
          page: 1,
          previousTokens: [],
          skipToken: undefined,
          filter: this.state.completedFilter
        }),
        this._loadAvailableVotes(),
        this._loadSuggestionTotals()
      ]);
    } finally {
      this._updateState({
        isActiveSuggestionsLoading: false,
        isCompletedSuggestionsLoading: false
      });
    }
  }

  private async _loadSuggestionTotals(): Promise<void> {
    try {
      const listId: string = this._getResolvedListId();
      const activeStatuses: string[] = this._getActiveStatuses();
      const completedStatus: string = this.state.completedStatus;

      const fetchCount = async (statuses: string[]): Promise<number | undefined> => {
        if (statuses.length === 0) {
          return 0;
        }

        const response = await this.props.graphService.getSuggestionItems(listId, {
          statuses,
          top: 999,
          orderBy: 'createdDateTime desc'
        });

        return typeof response.totalCount === 'number' ? response.totalCount : response.items.length;
      };

      const [activeSuggestionsTotal, completedSuggestionsTotal] = await Promise.all([
        fetchCount(activeStatuses),
        fetchCount([completedStatus])
      ]);

      this._updateState({ activeSuggestionsTotal, completedSuggestionsTotal });
    } catch (error) {
      console.error('Failed to load suggestion totals.', error);
    }
  }

  private async _loadAdminSuggestions(filter: IFilterState = this.state.adminFilter): Promise<void> {
    const resolvedCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      filter.category,
      this.state.categories
    );
    const resolvedSubcategory: string | undefined = this._normalizeFilterSubcategory(
      resolvedCategory,
      filter.subcategory,
      this.state.subcategories
    );

    const normalizedFilter: IFilterState = {
      ...filter,
      category: resolvedCategory,
      subcategory: resolvedSubcategory,
      status: this._normalizeAdminFilterStatus(filter.status)
    };

    this._updateState({
      isAdminSuggestionsLoading: true,
      error: undefined,
      success: undefined,
      adminFilter: normalizedFilter
    });

    try {
      const items: ISuggestionItem[] = await this._getTopSuggestionsByVotes(normalizedFilter);
      this._updateState({
        adminSuggestions: items,
        isAdminSuggestionsLoading: false
      });
    } catch (error) {
      this._handleError(strings.TopSuggestionsLoadErrorMessage, error);
      this._updateState({ isAdminSuggestionsLoading: false });
    }
  }

  private async _loadCategories(): Promise<void> {
    const listId: string = this._getResolvedCategoryListId();
    const itemsFromGraph: IGraphCategoryItem[] = await this.props.graphService.getCategoryItems(listId);

    const definitions: SuggestionCategory[] = itemsFromGraph
      .map((item) => {
        const rawTitle: unknown = item.fields?.Title;
        if (typeof rawTitle !== 'string') {
          return undefined;
        }

        const trimmed: string = rawTitle.trim();
        return trimmed.length > 0 ? trimmed : undefined;
      })
      .filter((value): value is SuggestionCategory => !!value);

    this._applyCategories(definitions);
  }

  private async _loadSubcategories(): Promise<void> {
    const listId: string = this._getResolvedSubcategoryListId();
    const itemsFromGraph: IGraphSubcategoryItem[] = await this.props.graphService.getSubcategoryItems(listId);

    const definitions: ISubcategoryDefinition[] = itemsFromGraph
      .map((item) => {
        const fields = item.fields ?? {};
        const rawTitle: unknown = (fields as { Title?: unknown }).Title;

        if (typeof rawTitle !== 'string') {
          return undefined;
        }

        const trimmedTitle: string = rawTitle.trim();

        if (!trimmedTitle) {
          return undefined;
        }

        const rawCategory: unknown = (fields as { Category?: unknown }).Category;
        const normalizedCategory: SuggestionCategory | undefined = this._tryNormalizeCategory(rawCategory);

        return {
          key: item.id.toString(),
          title: trimmedTitle,
          category: normalizedCategory
        } as ISubcategoryDefinition;
      })
      .filter((definition): definition is ISubcategoryDefinition => !!definition)
      .sort((a, b) => a.title.localeCompare(b.title));

    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      this.state.newCategory,
      this.state.newSubcategoryKey,
      definitions
    );

    const nextActiveFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      this.state.activeFilter.category,
      this.state.activeFilter.subcategory,
      definitions
    );

    const nextCompletedFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      this.state.completedFilter.category,
      this.state.completedFilter.subcategory,
      definitions
    );
    const nextAdminFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      this.state.adminFilter.category,
      this.state.adminFilter.subcategory,
      definitions
    );

    this._updateState(
      {
        subcategories: definitions,
        newSubcategoryKey: nextSubcategoryKey,
        activeFilter: { ...this.state.activeFilter, subcategory: nextActiveFilterSubcategory },
        completedFilter: { ...this.state.completedFilter, subcategory: nextCompletedFilterSubcategory },
        adminFilter: { ...this.state.adminFilter, subcategory: nextAdminFilterSubcategory }
      },
      () => {
        if (this.state.selectedMainTab === 'admin') {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._loadAdminSuggestions();
        }
      }
    );
  }

  private async _loadStatusesFromList(): Promise<void> {
    const listId: string = this._getResolvedStatusListId();
    const items = await this.props.graphService.getStatusItems(listId);

    const definitions: Array<{ title: string; order?: number; isCompleted: boolean }> = [];

    items.forEach((item) => {
      const rawTitle: unknown = item.fields?.Title;

      if (typeof rawTitle !== 'string') {
        return;
      }

      const title: string = rawTitle.trim();

      if (!title) {
        return;
      }

      const rawOrder: unknown = item.fields?.SortOrder;
      let order: number | undefined;

      if (typeof rawOrder === 'number' && Number.isFinite(rawOrder)) {
        order = rawOrder;
      } else if (typeof rawOrder === 'string') {
        const parsed: number = parseInt(rawOrder, 10);
        if (Number.isFinite(parsed)) {
          order = parsed;
        }
      }

      const isCompleted: boolean = this._parseBooleanValue(item.fields?.IsCompleted);

      definitions.push({ title, order, isCompleted });
    });

    definitions.sort((a, b) => {
      if (typeof a.order === 'number' && typeof b.order === 'number' && a.order !== b.order) {
        return a.order - b.order;
      }

      if (typeof a.order === 'number') {
        return -1;
      }

      if (typeof b.order === 'number') {
        return 1;
      }

      return a.title.localeCompare(b.title);
    });

    const statuses: string[] = definitions.map((definition) => definition.title);

    if (statuses.length === 0) {
      this._applyStatusConfiguration(undefined, undefined, { reloadSuggestions: false });
      return;
    }

    const completedStatusFromList: string | undefined = definitions.find(
      (definition) => definition.isCompleted
    )?.title;
    const completedStatus: string | undefined =
      completedStatusFromList ?? this.state.completedStatus;

    this._applyStatusConfiguration(statuses, completedStatus, { reloadSuggestions: false });
  }

  private _applyCategories(definitions: SuggestionCategory[]): void {
    const normalized: SuggestionCategory[] = this._normalizeCategoryList(definitions);
    const categories: SuggestionCategory[] = normalized.length > 0 ? normalized : [...FALLBACK_CATEGORIES];

    const nextCategory: SuggestionCategory =
      this._findCategoryMatch(this.state.newCategory, categories) ?? this._getDefaultCategory(categories);

    const nextActiveFilterCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      this.state.activeFilter.category,
      categories
    );
    const nextCompletedFilterCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      this.state.completedFilter.category,
      categories
    );
    const nextAdminFilterCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      this.state.adminFilter.category,
      categories
    );

    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      nextCategory,
      this.state.newSubcategoryKey,
      this.state.subcategories
    );

    const nextActiveFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      nextActiveFilterCategory,
      this.state.activeFilter.subcategory,
      this.state.subcategories
    );

    const nextCompletedFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      nextCompletedFilterCategory,
      this.state.completedFilter.subcategory,
      this.state.subcategories
    );
    const nextAdminFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      nextAdminFilterCategory,
      this.state.adminFilter.subcategory,
      this.state.subcategories
    );

    this._updateState(
      {
        categories,
        newCategory: nextCategory,
        newSubcategoryKey: nextSubcategoryKey,
        activeFilter: {
          ...this.state.activeFilter,
          category: nextActiveFilterCategory,
          subcategory: nextActiveFilterSubcategory
        },
        completedFilter: {
          ...this.state.completedFilter,
          category: nextCompletedFilterCategory,
          subcategory: nextCompletedFilterSubcategory
        },
        adminFilter: {
          ...this.state.adminFilter,
          category: nextAdminFilterCategory,
          subcategory: nextAdminFilterSubcategory
        }
      },
      () => {
        if (this.state.selectedMainTab === 'admin') {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._loadAdminSuggestions();
        }
      }
    );
  }

  private _normalizeCategoryList(values: SuggestionCategory[]): SuggestionCategory[] {
    const seen: Set<string> = new Set();
    const normalized: SuggestionCategory[] = [];

    values.forEach((value) => {
      const trimmed: string = value.trim();

      if (!trimmed) {
        return;
      }

      const key: string = this._getCategoryKey(trimmed);

      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      normalized.push(trimmed);
    });

    normalized.sort((a, b) => a.localeCompare(b));
    return normalized;
  }

  private _findCategoryMatch(
    value: SuggestionCategory | undefined,
    categories: SuggestionCategory[]
  ): SuggestionCategory | undefined {
    if (!value) {
      return undefined;
    }

    const normalized: string = value.trim();

    if (!normalized) {
      return undefined;
    }

    const lower: string = this._getCategoryKey(normalized);
    return categories.find((category) => category.toLowerCase() === lower);
  }

  private _getCategoryKey(category: SuggestionCategory): string {
    return category.trim().toLowerCase();
  }

  private _getVoteSummaryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    if (this.state.isUnlimitedVotes) {
      return [
        {
          key: 'unlimited',
          text: strings.VotesUnlimitedLabel,
          disabled: true
        }
      ];
    }

    const maxVotes: number = this._getMaxVotesPerCategory();
    const categoryKeys: Set<string> = new Set();
    categories.forEach((category) => categoryKeys.add(this._getCategoryKey(category)));
    Object.keys(this.state.availableVotesByCategory).forEach((key) => categoryKeys.add(key));

    if (categoryKeys.size === 0) {
      return [
        {
          key: 'default',
          text: strings.VotesPerCategoryDefaultLabel.replace('{0}', maxVotes.toString()),
          disabled: true
        }
      ];
    }

    const summaryEntries: IDropdownOption[] = [];

    categoryKeys.forEach((key) => {
      const displayName: string =
        categories.find((category) => this._getCategoryKey(category) === key) ?? key;
      const remaining: number = this.state.availableVotesByCategory[key] ?? maxVotes;
      summaryEntries.push({
        key,
        text: `${displayName}: ${remaining}/${maxVotes}`
      });
    });

    return summaryEntries;
  }

  private _getRemainingVotesForCategory(category: SuggestionCategory): number {
    if (this.state.isUnlimitedVotes) {
      return Number.POSITIVE_INFINITY;
    }

    const key: string = this._getCategoryKey(category);
    return this.state.availableVotesByCategory[key] ?? this._getMaxVotesPerCategory();
  }

  private _getDefaultCategory(categories: SuggestionCategory[]): SuggestionCategory {
    return categories[0] ?? DEFAULT_SUGGESTION_CATEGORY;
  }

  private _getMaxVotesPerCategory(): number {
    const value: number = this.props.totalVotesPerUser;

    if (!Number.isFinite(value) || value <= 0) {
      return DEFAULT_MAX_VOTES_PER_CATEGORY;
    }

    return Math.floor(value);
  }

  private _getTotalPages(totalCount: number | undefined, pageSize: number): number | undefined {
    if (typeof totalCount !== 'number' || !Number.isFinite(totalCount)) {
      return undefined;
    }

    if (pageSize <= 0 || !Number.isFinite(pageSize)) {
      return undefined;
    }

    return Math.max(1, Math.ceil(totalCount / Math.floor(pageSize)));
  }

  private async _fetchActiveSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    filter?: IFilterState;
  }): Promise<void> {
    const filter: IFilterState = options.filter ?? this.state.activeFilter;
    const hasSpecificSuggestion: boolean = typeof filter.suggestionId === 'number';
    const effectiveSkipToken: string | undefined = hasSpecificSuggestion ? undefined : options.skipToken;
    const effectivePreviousTokens: (string | undefined)[] = hasSpecificSuggestion
      ? []
      : options.previousTokens;

    const { items, nextToken, totalCount } = await this._getSuggestionsPage(
      'active',
      effectiveSkipToken,
      filter
    );
    const filteredItems: ISuggestionItem[] = this._filterDeniedSuggestions(items);

    if (!hasSpecificSuggestion && filteredItems.length === 0 && effectivePreviousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...effectivePreviousTokens];
      const previousToken: string | undefined = tokens.pop();

      await this._fetchActiveSuggestions({
        page: Math.max(options.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter
      });
      return;
    }

    this._updateState({
      activeSuggestions: {
        items: filteredItems,
        page: hasSpecificSuggestion ? 1 : options.page,
        currentToken: hasSpecificSuggestion ? undefined : effectiveSkipToken,
        nextToken: hasSpecificSuggestion ? undefined : nextToken,
        previousTokens: hasSpecificSuggestion ? [] : effectivePreviousTokens,
        totalCount: typeof totalCount === 'number' ? totalCount : filteredItems.length
      },
      activeFilter: filter,
      isActiveSuggestionsLoading: false
    });
  }

  private async _fetchCompletedSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    filter?: IFilterState;
  }): Promise<void> {
    const filter: IFilterState = options.filter ?? this.state.completedFilter;
    const hasSpecificSuggestion: boolean = typeof filter.suggestionId === 'number';
    const effectiveSkipToken: string | undefined = hasSpecificSuggestion ? undefined : options.skipToken;
    const effectivePreviousTokens: (string | undefined)[] = hasSpecificSuggestion
      ? []
      : options.previousTokens;

    const { items, nextToken, totalCount } = await this._getSuggestionsPage(
      'completed',
      effectiveSkipToken,
      filter
    );
    const filteredItems: ISuggestionItem[] =
      filter.includeDenied === true ? items : this._filterDeniedSuggestions(items);

    if (!hasSpecificSuggestion && filteredItems.length === 0 && effectivePreviousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...effectivePreviousTokens];
      const previousToken: string | undefined = tokens.pop();

      await this._fetchCompletedSuggestions({
        page: Math.max(options.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter
      });
      return;
    }

    this._updateState({
      completedSuggestions: {
        items: filteredItems,
        page: hasSpecificSuggestion ? 1 : options.page,
        currentToken: hasSpecificSuggestion ? undefined : effectiveSkipToken,
        nextToken: hasSpecificSuggestion ? undefined : nextToken,
        previousTokens: hasSpecificSuggestion ? [] : effectivePreviousTokens,
        totalCount: typeof totalCount === 'number' ? totalCount : filteredItems.length
      },
      completedFilter: filter,
      isCompletedSuggestionsLoading: false
    });
  }

  private async _refreshActiveSuggestions(): Promise<void> {
    const { activeSuggestions, activeFilter } = this.state;

    this._updateState({ isActiveSuggestionsLoading: true });

    try {
      await this._fetchActiveSuggestions({
        page: activeSuggestions.page,
        previousTokens: activeSuggestions.previousTokens,
        skipToken: activeSuggestions.currentToken,
        filter: activeFilter
      });
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  }

  private async _refreshCompletedSuggestions(): Promise<void> {
    const { completedSuggestions, completedFilter } = this.state;

    this._updateState({ isCompletedSuggestionsLoading: true });

    try {
      await this._fetchCompletedSuggestions({
        page: completedSuggestions.page,
        previousTokens: completedSuggestions.previousTokens,
        skipToken: completedSuggestions.currentToken,
        filter: completedFilter
      });
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  }

  private async _getSuggestionsPage(
    statusGroup: 'active' | 'completed',
    skipToken: string | undefined,
    filter: IFilterState
  ): Promise<{ items: ISuggestionItem[]; nextToken?: string; totalCount?: number }> {
    const listId: string = this._getResolvedListId();
    const baseStatuses: string[] =
      statusGroup === 'completed' ? this._getCompletedStatuses(filter) : this._getActiveStatuses();
    const normalizedStatus: string | undefined = this._normalizeStatusValue(
      filter.status,
      this.state.statuses
    );
    const hasSpecificSuggestion: boolean = typeof filter.suggestionId === 'number';
    const allowStatusOverride: boolean = statusGroup !== 'completed' || filter.includeDenied !== true;
    const statuses: string[] =
      normalizedStatus &&
      (hasSpecificSuggestion || (allowStatusOverride && this._isStatusInCollection(normalizedStatus, baseStatuses)))
        ? [normalizedStatus]
        : baseStatuses;

    const pageSize: number =
      statusGroup === 'completed' ? this.state.completedPageSize : this.state.activePageSize;

    const orderBy: string =
      statusGroup === 'completed' ? 'fields/CompletedDateTime desc' : 'createdDateTime desc';

    const response = await this.props.graphService.getSuggestionItems(listId, {
      statuses,
      top: pageSize,
      skipToken,
      category: filter.category,
      subcategory: filter.subcategory,
      searchQuery: filter.searchQuery,
      suggestionIds:
        typeof filter.suggestionId === 'number' ? [filter.suggestionId] : undefined,
      orderBy
    });

    const suggestionIds: number[] = response.items
      .map((entry) =>
        this._parseNumericId(entry.listItemId ?? entry.fields.id ?? (entry.fields as { Id?: unknown }).Id)
      )
      .filter((value): value is number => typeof value === 'number');

    let votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    const shouldLoadVotes: boolean =
      suggestionIds.length > 0 &&
      statuses.some(
        (status) =>
          !this._isCompletedStatusValue(
            status,
            this.state.completedStatus,
            this.state.deniedStatus
          )
      );

    if (shouldLoadVotes) {
      const voteListId: string = this._getResolvedVotesListId();
      const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
        suggestionIds
      });
      votesBySuggestion = this._groupVotesBySuggestion(voteItems);
    }

    let commentCounts: Map<number, number> = new Map();

    if (suggestionIds.length > 0 && this._currentCommentsListId) {
      const commentListId: string = this._getResolvedCommentsListId();
      commentCounts = await this.props.graphService.getCommentCounts(commentListId, {
        suggestionIds
      });
    }

    const items: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
      response.items,
      votesBySuggestion,
      commentCounts
    );

    let mismatchIndex: number = -1;
    if (statusGroup === 'active' && items.length > 1) {
      mismatchIndex = items.findIndex((item, index) => {
        if (index === 0) {
          return false;
        }

        return (
          getSortableDateValue(items[index - 1].createdDateTime) <
          getSortableDateValue(item.createdDateTime)
        );
      });
    }

    const shouldForceSort: boolean = isClientSortForced();
    const shouldSortClientSide: boolean = statusGroup === 'active' && (mismatchIndex !== -1 || shouldForceSort);
    const sortedItems: ISuggestionItem[] = shouldSortClientSide
      ? [...items].sort((a, b) => {
          const timeDelta: number =
            getSortableDateValue(b.createdDateTime) - getSortableDateValue(a.createdDateTime);
          if (timeDelta !== 0) {
            return timeDelta;
          }

          return b.id - a.id;
        })
      : items;

    if (statusGroup === 'active' && isSortDiagnosticsEnabled()) {
      if (mismatchIndex === -1 && !shouldForceSort) {
        console.info('Active suggestions sorted by createdDateTime desc.', {
          orderBy,
          first: items[0]?.createdDateTime,
          last: items[items.length - 1]?.createdDateTime
        });
      } else {
        console.warn('Active suggestions not sorted by createdDateTime desc. Applying client sort.', {
          orderBy,
          mismatchIndex,
          previous: items[mismatchIndex - 1]?.createdDateTime,
          current: items[mismatchIndex]?.createdDateTime
        });
      }
    }

    return { items: sortedItems, nextToken: response.nextToken, totalCount: response.totalCount };
  }

  private async _getTopSuggestionsByVotes(filter: IFilterState): Promise<ISuggestionItem[]> {
    const listId: string = this._getResolvedListId();
    const activeStatuses: string[] = this._getActiveStatuses();
    const filterStatus: string | undefined = this._normalizeAdminFilterStatus(filter.status);
    const statuses: string[] =
      filterStatus && this._isStatusInCollection(filterStatus, activeStatuses)
        ? [filterStatus]
        : activeStatuses;

    if (statuses.length === 0) {
      return [];
    }

    const response = await this.props.graphService.getSuggestionItems(listId, {
      statuses,
      top: ADMIN_TOP_SUGGESTIONS_COUNT,
      category: filter.category,
      subcategory: filter.subcategory,
      orderBy: 'fields/Votes desc'
    });

    const suggestionIds: number[] = response.items
      .map((entry) =>
        this._parseNumericId(entry.listItemId ?? entry.fields.id ?? (entry.fields as { Id?: unknown }).Id)
      )
      .filter((value): value is number => typeof value === 'number');

    let votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    if (suggestionIds.length > 0) {
      const voteListId: string = this._getResolvedVotesListId();
      const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
        suggestionIds
      });
      votesBySuggestion = this._groupVotesBySuggestion(voteItems);
    }

    let commentCounts: Map<number, number> = new Map();

    if (suggestionIds.length > 0 && this._currentCommentsListId) {
      const commentListId: string = this._getResolvedCommentsListId();
      commentCounts = await this.props.graphService.getCommentCounts(commentListId, {
        suggestionIds
      });
    }

    const suggestions = this._mapGraphItemsToSuggestions(response.items, votesBySuggestion, commentCounts);

    const filteredSuggestions: ISuggestionItem[] = suggestions.filter((item) => {
      const isCompleted: boolean = this._isCompletedStatusValue(
        item.status,
        this.state.completedStatus,
        this.state.deniedStatus
      );

      return item.votes > 0 && !isCompleted;
    });

    return [...filteredSuggestions].sort((left, right) => {
      const voteDelta: number = right.votes - left.votes;
      if (voteDelta !== 0) {
        return voteDelta;
      }

      const dateDelta: number =
        getSortableDateValue(right.createdDateTime) - getSortableDateValue(left.createdDateTime);
      if (dateDelta !== 0) {
        return dateDelta;
      }

      return right.id - left.id;
    });
  }

  private _mapGraphItemsToSuggestions(
    graphItems: IGraphSuggestionItem[],
    votesBySuggestion: Map<number, IVoteEntry[]>,
    commentCounts: Map<number, number>
  ): ISuggestionItem[] {
    return graphItems
      .map((entry) => {
        const fields: IGraphSuggestionItemFields = entry.fields;
        const rawId: unknown = entry.listItemId ?? fields.id ?? (fields as { Id?: unknown }).Id;
        const suggestionId: number | undefined = this._parseNumericId(rawId);

        if (typeof suggestionId !== 'number') {
          return undefined;
        }

        const voteEntries: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];
        const storedVotes: number = this._parseVotes(fields.Votes);
        const rawStatus: unknown = fields.Status;
        const resolvedStatus: string =
          typeof rawStatus === 'string' && rawStatus.trim().length > 0
            ? rawStatus.trim()
            : this.state.defaultStatus;
        const isCompleted: boolean = this._isCompletedStatusValue(
          resolvedStatus,
          this.state.completedStatus,
          this.state.deniedStatus
        );
        const liveVotes: number = voteEntries.reduce((total, vote) => total + vote.votes, 0);
        const votes: number = isCompleted ? Math.max(liveVotes, storedVotes) : liveVotes;
        const createdDateTimeFromFields: string | undefined =
          typeof fields.Created === 'string' && fields.Created.trim().length > 0
            ? fields.Created.trim()
            : undefined;
        const createdDateTimeFromEntry: string | undefined =
          typeof entry.createdDateTime === 'string' && entry.createdDateTime.trim().length > 0
            ? entry.createdDateTime.trim()
            : undefined;
        const createdDateTime: string | undefined = createdDateTimeFromFields ?? createdDateTimeFromEntry;
        const lastModifiedDateTime: string | undefined =
          typeof entry.lastModifiedDateTime === 'string' && entry.lastModifiedDateTime.trim().length > 0
            ? entry.lastModifiedDateTime.trim()
            : undefined;
        const completedDateTime: string | undefined =
          typeof fields.CompletedDateTime === 'string' && fields.CompletedDateTime.trim().length > 0
            ? fields.CompletedDateTime.trim()
            : undefined;
        const commentCount: number = commentCounts.get(suggestionId) ?? 0;

        return {
          id: suggestionId,
          title:
            typeof fields.Title === 'string' && fields.Title.trim().length > 0
              ? fields.Title
              : 'Untitled suggestion',
          description: typeof fields.Details === 'string' ? fields.Details : '',
          votes,
          status: resolvedStatus,
          category: this._normalizeCategory(fields.Category),
          subcategory:
            typeof fields.Subcategory === 'string' && fields.Subcategory.trim().length > 0
              ? fields.Subcategory.trim()
              : undefined,
          voters: voteEntries.map((vote) => vote.username),
          createdByLoginName: this._normalizeLoginName(entry.createdByUserPrincipalName),
          createdDateTime,
          lastModifiedDateTime,
          completedDateTime,
          voteEntries,
          commentCount,
          comments: [],
          areCommentsLoaded: commentCount === 0
        } as ISuggestionItem;
      })
      .filter((item): item is ISuggestionItem => !!item);
  }

  private _groupVotesBySuggestion(voteItems: IGraphVoteItem[]): Map<number, IVoteEntry[]> {
    const votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    voteItems.forEach((entry: IGraphVoteItem) => {
      const fields = entry.fields ?? {};
      const suggestionId: number | undefined = this._parseNumericId(
        (fields as { SuggestionId?: unknown }).SuggestionId
      );
      const rawUsername: unknown = (fields as { Username?: unknown }).Username;
      const normalizedUsername: string | undefined = this._normalizeLoginName(
        typeof rawUsername === 'string' ? rawUsername : undefined
      );
      const votes: number = this._parseVotes((fields as { Votes?: unknown }).Votes);

      if (!suggestionId || !normalizedUsername || votes <= 0) {
        return;
      }

      const entriesForSuggestion: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];
      entriesForSuggestion.push({
        id: entry.id,
        username: normalizedUsername,
        votes
      });
      votesBySuggestion.set(suggestionId, entriesForSuggestion);
    });

    return votesBySuggestion;
  }

  private _groupCommentsBySuggestion(commentItems: IGraphCommentItem[]): Map<number, ISuggestionComment[]> {
    const commentsBySuggestion: Map<number, ISuggestionComment[]> = new Map();

    commentItems.forEach((entry) => {
      const fields = entry.fields ?? {};
      const suggestionId: number | undefined = this._parseNumericId(
        (fields as { SuggestionId?: unknown }).SuggestionId
      );
      const rawComment: unknown =
        (fields as { Comment?: unknown }).Comment ?? (fields as { Title?: unknown }).Title;
      const commentText: string | undefined = this._normalizeCommentValue(rawComment);

      if (!suggestionId || !commentText) {
        return;
      }

      const createdDateTime: string | undefined =
        typeof entry.createdDateTime === 'string' && entry.createdDateTime.trim().length > 0
          ? entry.createdDateTime.trim()
          : undefined;

      const displayName: string | undefined =
        typeof entry.createdByUserDisplayName === 'string' && entry.createdByUserDisplayName.trim().length > 0
          ? entry.createdByUserDisplayName.trim()
          : undefined;
      const principalName: string | undefined =
        typeof entry.createdByUserPrincipalName === 'string' && entry.createdByUserPrincipalName.trim().length > 0
          ? entry.createdByUserPrincipalName.trim()
          : undefined;
      const author: string | undefined = displayName ?? principalName;

      const existing: ISuggestionComment[] = commentsBySuggestion.get(suggestionId) ?? [];
      existing.push({
        id: entry.id,
        text: commentText,
        author,
        createdDateTime
      });
      commentsBySuggestion.set(suggestionId, existing);
    });

    commentsBySuggestion.forEach((comments, key) => {
      const sorted: ISuggestionComment[] = [...comments].sort((a, b) => {
        if (!a.createdDateTime && !b.createdDateTime) {
          return a.id - b.id;
        }

        if (!a.createdDateTime) {
          return -1;
        }

        if (!b.createdDateTime) {
          return 1;
        }

        return new Date(a.createdDateTime).getTime() - new Date(b.createdDateTime).getTime();
      });
      commentsBySuggestion.set(key, sorted);
    });

    return commentsBySuggestion;
  }

  private _normalizeCommentValue(value: unknown): string | undefined {
    if (typeof value === 'string') {
      const trimmed: string = value.trim();
      return trimmed.length > 0 ? trimmed : undefined;
    }

    if (value && typeof value === 'object') {
      const richText: unknown = (value as { $content?: unknown }).$content;

      if (typeof richText === 'string') {
        const trimmed: string = richText.trim();
        return trimmed.length > 0 ? trimmed : undefined;
      }
    }

    return undefined;
  }

  private async _loadAvailableVotes(): Promise<void> {
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    if (!normalizedUser) {
      this._updateState({
        availableVotesByCategory: {},
        myVoteSuggestions: [],
        isMyVotesLoading: false
      });
      return;
    }

    const listId: string = this._getResolvedListId();
    const voteListId: string = this._getResolvedVotesListId();
    this._updateState({ isMyVotesLoading: true });

    try {
      const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
        username: normalizedUser
      });

      const votedSuggestionIds: number[] = voteItems
        .map((entry) => this._parseNumericId(entry.fields?.SuggestionId))
        .filter((value): value is number => typeof value === 'number');

      const emptySuggestionResponse: { items: IGraphSuggestionItem[] } = { items: [] };
      const [votedSuggestionsResponse, createdSuggestionsResponse] = await Promise.all([
        votedSuggestionIds.length > 0
          ? this.props.graphService.getSuggestionItems(listId, {
              suggestionIds: votedSuggestionIds,
              top: votedSuggestionIds.length
            })
          : Promise.resolve(emptySuggestionResponse),
        this.props.graphService.getSuggestionItems(listId, {
          createdByUserPrincipalName: normalizedUser
        })
      ]);

      const combinedGraphItems: IGraphSuggestionItem[] = [
        ...votedSuggestionsResponse.items,
        ...createdSuggestionsResponse.items
      ];

      const suggestionIdSet: Set<number> = new Set();
      combinedGraphItems.forEach((entry) => {
        const suggestionId: number | undefined = this._parseNumericId(
          entry.listItemId ?? entry.fields.id ?? (entry.fields as { Id?: unknown }).Id
        );

        if (typeof suggestionId === 'number') {
          suggestionIdSet.add(suggestionId);
        }
      });

      const suggestionIds: number[] = [];
      suggestionIdSet.forEach((value) => {
        suggestionIds.push(value);
      });

      if (suggestionIds.length === 0) {
        this._updateState({
          availableVotesByCategory: {},
          myVoteSuggestions: [],
          isMyVotesLoading: false
        });
        return;
      }

      const [allVoteItems, commentCounts] = await Promise.all([
        this.props.graphService.getVoteItems(voteListId, {
          suggestionIds
        }),
        this._currentCommentsListId
          ? this.props.graphService.getCommentCounts(this._getResolvedCommentsListId(), {
              suggestionIds
            })
          : Promise.resolve(new Map<number, number>())
      ]);

      const votesBySuggestion: Map<number, IVoteEntry[]> = this._groupVotesBySuggestion(allVoteItems);
      const mappedSuggestions: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
        combinedGraphItems,
        votesBySuggestion,
        commentCounts
      );

      const suggestionsById: Map<number, ISuggestionItem> = new Map();
      mappedSuggestions.forEach((suggestion) => {
        suggestionsById.set(suggestion.id, suggestion);
      });

      const suggestions: ISuggestionItem[] = [];
      suggestionsById.forEach((suggestion) => {
        suggestions.push(suggestion);
      });

      const usedVotesByCategory: Record<string, number> = {};

      voteItems.forEach((entry) => {
        const suggestionId: number | undefined = this._parseNumericId(entry.fields?.SuggestionId);

        if (typeof suggestionId !== 'number') {
          return;
        }

        const suggestion: ISuggestionItem | undefined = suggestionsById.get(suggestionId);

        if (!suggestion) {
          return;
        }

        const votes: number = this._parseVotes(entry.fields?.Votes);

        if (votes <= 0) {
          return;
        }

        const categoryKey: string = this._getCategoryKey(suggestion.category);
        usedVotesByCategory[categoryKey] = (usedVotesByCategory[categoryKey] ?? 0) + votes;
      });

      const availableVotesByCategory: Record<string, number> = {};
      const maxVotes: number = this._getMaxVotesPerCategory();
      Object.keys(usedVotesByCategory).forEach((key) => {
        const remaining: number = Math.max(maxVotes - usedVotesByCategory[key], 0);
        availableVotesByCategory[key] = remaining;
      });

      this._updateState({
        availableVotesByCategory,
        myVoteSuggestions: suggestions,
        isMyVotesLoading: false
      });
    } catch (error) {
      console.error('Failed to load available votes.', error);
      this._updateState({ isMyVotesLoading: false });
    }
  }

  private _onTitleChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._updateState({ newTitle: newValue ?? '' }, () => {
      this._handleSimilarSuggestionsInput(this.state.newTitle, this.state.newDescription);
    });
  };

  private _onDescriptionEditorChange = (value: string): void => {
    this._setDescriptionValue(value);
  };

  private _setDescriptionValue(value: string): void {
    this._updateState({ newDescription: value }, () => {
      this._handleSimilarSuggestionsInput(this.state.newTitle, value);
    });
  }

  private _areSimilarSuggestionQueriesEqual(
    left: ISimilarSuggestionsQuery,
    right: ISimilarSuggestionsQuery
  ): boolean {
    return left.title === right.title && left.description === right.description;
  }

  private _handleSimilarSuggestionsInput(title: string, description: string): void {
    const normalizedTitle: string = (title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = getPlainTextFromHtml(description ?? '');
    const hasTitleQuery: boolean = normalizedTitle.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH;
    const hasDescriptionQuery: boolean =
      normalizedDescription.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH;

    if (!hasTitleQuery && !hasDescriptionQuery) {
      this._debouncedSimilarSuggestionsSearch.cancel();
      const previousSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const nextExpandedCommentIds: number[] =
        typeof previousSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== previousSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        typeof previousSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== previousSelectedId)
          : this.state.loadingCommentIds;
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
      this._pendingSimilarSuggestionsQuery = undefined;
      return;
    }

    const nextQuery: ISimilarSuggestionsQuery = {
      title: hasTitleQuery ? normalizedTitle : '',
      description: hasDescriptionQuery ? normalizedDescription : ''
    };

    if (
      this._areSimilarSuggestionQueriesEqual(nextQuery, this.state.similarSuggestionsQuery) &&
      !this.state.isSimilarSuggestionsLoading &&
      this.state.similarSuggestions.items.length > 0
    ) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = nextQuery;

    if (!this._currentListId) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = undefined;
    this._debouncedSimilarSuggestionsSearch(nextQuery);
  }

  private async _searchSimilarSuggestions(query: ISimilarSuggestionsQuery): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = undefined;

    const normalizedTitle: string = (query.title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = (query.description ?? '').replace(/\s+/g, ' ').trim();
    const effectiveQuery: ISimilarSuggestionsQuery = {
      title:
        normalizedTitle.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH ? normalizedTitle : '',
      description:
        normalizedDescription.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH
          ? normalizedDescription
          : ''
    };

    if (!effectiveQuery.title && !effectiveQuery.description) {
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false
      });
      return;
    }

    await this._fetchSimilarSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      query: effectiveQuery
    });
  }

  private async _fetchSimilarSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    query: ISimilarSuggestionsQuery;
  }): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    const listId: string | undefined = this._currentListId;

    if (!listId) {
      return;
    }

    const normalizedTitle: string = (options.query.title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = (options.query.description ?? '').replace(/\s+/g, ' ').trim();
    const effectiveQuery: ISimilarSuggestionsQuery = {
      title: normalizedTitle,
      description: normalizedDescription
    };

    if (!effectiveQuery.title && !effectiveQuery.description) {
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false
      });
      return;
    }

    this._updateState({
      similarSuggestionsQuery: { ...effectiveQuery },
      isSimilarSuggestionsLoading: true
    });

    try {
      const response = await this.props.graphService.getSuggestionItems(listId, {
        top: MAX_SIMILAR_SUGGESTIONS,
        skipToken: options.skipToken,
        titleSearchQuery: effectiveQuery.title || undefined,
        descriptionSearchQuery: effectiveQuery.description || undefined,
        orderBy: 'createdDateTime desc'
      });

      const suggestionIds: number[] = response.items
        .map((entry) => {
          const fields: IGraphSuggestionItemFields = entry.fields;
          const rawId: unknown = entry.listItemId ?? fields.id ?? (fields as { Id?: unknown }).Id;
          return this._parseNumericId(rawId);
        })
        .filter((value): value is number => typeof value === 'number');

      let votesBySuggestion: Map<number, IVoteEntry[]> = new Map();
      let commentCounts: Map<number, number> = new Map();
      let commentsBySuggestion: Map<number, ISuggestionComment[]> = new Map();

      if (suggestionIds.length > 0) {
        if (this._currentVotesListId) {
          const voteListId: string = this._getResolvedVotesListId();
          const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(
            voteListId,
            { suggestionIds }
          );
          votesBySuggestion = this._groupVotesBySuggestion(voteItems);
        }

        if (this._currentCommentsListId) {
          const commentListId: string = this._getResolvedCommentsListId();
          const [counts, commentItems] = await Promise.all([
            this.props.graphService.getCommentCounts(commentListId, { suggestionIds }),
            this.props.graphService.getCommentItems(commentListId, { suggestionIds })
          ]);
          commentCounts = counts;
          commentsBySuggestion = this._groupCommentsBySuggestion(commentItems);
        }
      }

      const baseItems: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
        response.items,
        votesBySuggestion,
        commentCounts
      );

      const enrichedItems: ISuggestionItem[] = this._currentCommentsListId
        ? baseItems.map((item) => {
            const loadedComments: ISuggestionComment[] = commentsBySuggestion.get(item.id) ?? [];
            const mappedCount: number = Math.max(
              loadedComments.length,
              commentCounts.get(item.id) ?? 0
            );

            return {
              ...item,
              comments: loadedComments,
              commentCount: mappedCount,
              areCommentsLoaded: true
            };
          })
        : baseItems;

      if (
        !this._areSimilarSuggestionQueriesEqual(this.state.similarSuggestionsQuery, effectiveQuery)
      ) {
        return;
      }

      const filteredItems: ISuggestionItem[] = this._filterDeniedSuggestions(enrichedItems);
      const limited: ISuggestionItem[] = filteredItems.slice(0, MAX_SIMILAR_SUGGESTIONS);

      const currentSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const shouldKeepSelection: boolean =
        typeof currentSelectedId === 'number' &&
        limited.some((entry) => entry.id === currentSelectedId);
      const nextExpandedCommentIds: number[] =
        !shouldKeepSelection && typeof currentSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== currentSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        !shouldKeepSelection && typeof currentSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== currentSelectedId)
          : this.state.loadingCommentIds;

      const nextSelectedSimilarSuggestion: ISuggestionItem | undefined = shouldKeepSelection
        ? limited.find((entry) => entry.id === currentSelectedId) ?? this.state.selectedSimilarSuggestion
        : undefined;

      const nextIsSelectedSimilarSuggestionLoading: boolean = shouldKeepSelection
        ? this.state.isSelectedSimilarSuggestionLoading
        : false;

      this._updateState({
        similarSuggestions: {
          items: limited,
          page: options.page,
          currentToken: options.skipToken,
          nextToken: response.nextToken,
          previousTokens: options.previousTokens
        },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: nextSelectedSimilarSuggestion,
        isSelectedSimilarSuggestionLoading: nextIsSelectedSimilarSuggestionLoading,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
    } catch (error) {
      console.error('Failed to load similar suggestions.', error);

      if (
        !this._areSimilarSuggestionQueriesEqual(this.state.similarSuggestionsQuery, effectiveQuery)
      ) {
        return;
      }

      const staleSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const nextExpandedCommentIds: number[] =
        typeof staleSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== staleSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        typeof staleSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== staleSelectedId)
          : this.state.loadingCommentIds;

      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
    }
  }

  private async _loadSelectedSimilarSuggestion(
    suggestionId: number,
    status: string
  ): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    this._updateState({ isSelectedSimilarSuggestionLoading: true });

    try {
      const isCompleted: boolean = this._isCompletedStatusValue(
        status,
        this.state.completedStatus,
        this.state.deniedStatus
      );
      const { items } = await this._getSuggestionsPage(isCompleted ? 'completed' : 'active', undefined, {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId,
        status
      });

      if (!this._isMounted || this.state.selectedSimilarSuggestion?.id !== suggestionId) {
        return;
      }

      const nextSuggestion: ISuggestionItem | undefined = items.find((entry) => entry.id === suggestionId);

      if (!nextSuggestion) {
        const nextExpanded: number[] = this.state.expandedCommentIds.filter((id) => id !== suggestionId);
        const nextLoading: number[] = this.state.loadingCommentIds.filter((id) => id !== suggestionId);

        this._updateState({
          selectedSimilarSuggestion: undefined,
          isSelectedSimilarSuggestionLoading: false,
          expandedCommentIds: nextExpanded,
          loadingCommentIds: nextLoading
        });
        return;
      }

      let comments: ISuggestionComment[] = [];
      let areCommentsLoaded: boolean = nextSuggestion.commentCount === 0;
      let commentCount: number = nextSuggestion.commentCount;

      if (nextSuggestion.commentCount > 0) {
        const commentListId: string = this._getResolvedCommentsListId();
        const commentItems: IGraphCommentItem[] = await this.props.graphService.getCommentItems(
          commentListId,
          {
            suggestionIds: [suggestionId]
          }
        );
        const commentsBySuggestion: Map<number, ISuggestionComment[]> = this._groupCommentsBySuggestion(
          commentItems
        );
        comments = commentsBySuggestion.get(suggestionId) ?? [];
        commentCount = comments.length;
        areCommentsLoaded = true;
      }

      this._updateState(
        {
          selectedSimilarSuggestion: {
            ...nextSuggestion,
            comments,
            commentCount,
            areCommentsLoaded
          },
          isSelectedSimilarSuggestionLoading: false
        },
        () => {
          this._ensureCommentSectionExpanded(suggestionId);
        }
      );
    } catch (error) {
      if (!this._isMounted || this.state.selectedSimilarSuggestion?.id !== suggestionId) {
        return;
      }

      this._handleError(strings.SelectedSuggestionLoadErrorMessage, error);

      const nextExpanded: number[] = this.state.expandedCommentIds.filter((id) => id !== suggestionId);
      const nextLoading: number[] = this.state.loadingCommentIds.filter((id) => id !== suggestionId);

      this._updateState({
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpanded,
        loadingCommentIds: nextLoading
      });
    }
  }

  private _onActiveSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = {
      ...this.state.activeFilter,
      searchQuery: newValue ?? '',
      suggestionId: undefined
    };
    this._updateState({ activeFilter: nextFilter });
    this._debouncedActiveFilterSearch(nextFilter);
  };

  private _onCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const normalized: string = key.trim();
    const nextCategory: SuggestionCategory =
      this._findCategoryMatch(normalized, this.state.categories) ?? this._getDefaultCategory(this.state.categories);
    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      nextCategory,
      this.state.newSubcategoryKey
    );

    this._updateState({ newCategory: nextCategory, newSubcategoryKey: nextSubcategoryKey });
  };

  private _onSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const definition: ISubcategoryDefinition | undefined = this.state.subcategories.find(
      (item) => item.key === key
    );

    if (!definition) {
      return;
    }

    this._updateState({ newSubcategoryKey: definition.key });
  };

  private _onActiveFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();
    let nextCategory: SuggestionCategory | undefined;

    if (key !== ALL_CATEGORY_FILTER_KEY) {
      nextCategory =
        this._findCategoryMatch(key, this.state.categories) ?? (key.length > 0 ? key : undefined);
    }

    const nextFilter: IFilterState = {
      ...this.state.activeFilter,
      category: nextCategory,
      subcategory: this._normalizeFilterSubcategory(
        nextCategory,
        this.state.activeFilter.subcategory,
        this.state.subcategories
      ),
      suggestionId: undefined
    };
    this._debouncedActiveFilterSearch.cancel();
    this._applyActiveFilter(nextFilter);
  };

  private _onActiveFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_SUBCATEGORY_FILTER_KEY
        ? { ...this.state.activeFilter, subcategory: undefined, suggestionId: undefined }
        : { ...this.state.activeFilter, subcategory: key, suggestionId: undefined };
    this._debouncedActiveFilterSearch.cancel();
    this._applyActiveFilter(nextFilter);
  };

  private _onActivePageSizeChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const nextSize: number | undefined = this._parsePageSizeOption(option);

    if (typeof nextSize !== 'number' || nextSize === this.state.activePageSize) {
      return;
    }

    const nextFilter: IFilterState = { ...this.state.activeFilter };

    this._updateState({ activePageSize: nextSize }, () => {
      this._applyActiveFilter(nextFilter);
    });
  };

  private _onCompletedSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = {
      ...this.state.completedFilter,
      searchQuery: newValue ?? '',
      suggestionId: undefined
    };
    this._updateState({ completedFilter: nextFilter });
    this._debouncedCompletedFilterSearch(nextFilter);
  };

  private _onCompletedFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();
    let nextCategory: SuggestionCategory | undefined;

    if (key !== ALL_CATEGORY_FILTER_KEY) {
      nextCategory =
        this._findCategoryMatch(key, this.state.categories) ?? (key.length > 0 ? key : undefined);
    }

    const nextFilter: IFilterState = {
      ...this.state.completedFilter,
      category: nextCategory,
      subcategory: this._normalizeFilterSubcategory(
        nextCategory,
        this.state.completedFilter.subcategory,
        this.state.subcategories
      ),
      suggestionId: undefined
    };
    this._debouncedCompletedFilterSearch.cancel();
    this._applyCompletedFilter(nextFilter);
  };

  private _onCompletedPageSizeChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    const nextSize: number | undefined = this._parsePageSizeOption(option);

    if (typeof nextSize !== 'number' || nextSize === this.state.completedPageSize) {
      return;
    }

    const nextFilter: IFilterState = { ...this.state.completedFilter };

    this._updateState({ completedPageSize: nextSize }, () => {
      this._applyCompletedFilter(nextFilter);
    });
  };

  private _onCompletedFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_SUBCATEGORY_FILTER_KEY
        ? { ...this.state.completedFilter, subcategory: undefined, suggestionId: undefined }
        : { ...this.state.completedFilter, subcategory: key, suggestionId: undefined };
    this._debouncedCompletedFilterSearch.cancel();
    this._applyCompletedFilter(nextFilter);
  };

  private _onCompletedDeniedFilterChange = (
    _event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    const nextIncludeDenied: boolean = checked === true;

    if (nextIncludeDenied === this.state.completedFilter.includeDenied) {
      return;
    }

    const nextFilter: IFilterState = {
      ...this.state.completedFilter,
      includeDenied: nextIncludeDenied,
      suggestionId: undefined
    };

    this._applyCompletedFilter(nextFilter);
  };

  private _onAdminFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();
    let nextCategory: SuggestionCategory | undefined;

    if (key !== ALL_CATEGORY_FILTER_KEY) {
      nextCategory =
        this._findCategoryMatch(key, this.state.categories) ?? (key.length > 0 ? key : undefined);
    }

    const nextFilter: IFilterState = {
      ...this.state.adminFilter,
      category: nextCategory,
      subcategory: this._normalizeFilterSubcategory(
        nextCategory,
        this.state.adminFilter.subcategory,
        this.state.subcategories
      ),
      suggestionId: undefined
    };

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._loadAdminSuggestions(nextFilter);
  };

  private _onAdminFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_SUBCATEGORY_FILTER_KEY
        ? { ...this.state.adminFilter, subcategory: undefined, suggestionId: undefined }
        : { ...this.state.adminFilter, subcategory: key, suggestionId: undefined };

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._loadAdminSuggestions(nextFilter);
  };

  private _onAdminFilterStatusChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_STATUS_FILTER_KEY
        ? { ...this.state.adminFilter, status: undefined, suggestionId: undefined }
        : { ...this.state.adminFilter, status: key, suggestionId: undefined };

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._loadAdminSuggestions(nextFilter);
  };

  private _getDefaultActiveFilter(): IFilterState {
    return {
      searchQuery: '',
      category: undefined,
      subcategory: undefined,
      suggestionId: undefined,
      status: undefined
    };
  }

  private _getDefaultCompletedFilter(): IFilterState {
    return {
      searchQuery: '',
      category: undefined,
      subcategory: undefined,
      suggestionId: undefined,
      status: this.state.completedStatus,
      includeDenied: false
    };
  }

  private _getDefaultAdminFilter(): IFilterState {
    return {
      searchQuery: '',
      category: undefined,
      subcategory: undefined,
      suggestionId: undefined,
      status: undefined
    };
  }

  private _hasSearchFilters(filter: IFilterState): boolean {
    return (
      (filter.searchQuery ?? '').trim().length > 0 ||
      !!filter.category ||
      !!filter.subcategory ||
      typeof filter.suggestionId === 'number' ||
      filter.includeDenied === true
    );
  }

  private _hasAdminFilters(filter: IFilterState): boolean {
    return this._hasSearchFilters(filter) || !!filter.status;
  }

  private _clearActiveFilters = (): void => {
    if (!this._hasSearchFilters(this.state.activeFilter)) {
      return;
    }
    this._debouncedActiveFilterSearch.cancel();
    this._applyActiveFilter(this._getDefaultActiveFilter());
  };

  private _clearCompletedFilters = (): void => {
    if (!this._hasSearchFilters(this.state.completedFilter)) {
      return;
    }
    this._debouncedCompletedFilterSearch.cancel();
    this._applyCompletedFilter(this._getDefaultCompletedFilter());
  };

  private _clearAdminFilters = (): void => {
    if (!this._hasAdminFilters(this.state.adminFilter)) {
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._loadAdminSuggestions(this._getDefaultAdminFilter());
  };

  private _dismissError = (): void => {
    this._updateState({ error: undefined });
  };

  private _dismissSuccess = (): void => {
    this._updateState({ success: undefined });
  };

  private _applyActiveFilter(nextFilter: IFilterState): void {
    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    this._fetchActiveSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      filter: nextFilter
    })
      .then(() => {
        this._updateState({ isActiveSuggestionsLoading: false });
      })
      .catch((error) => {
        this._handleError(strings.ActiveSuggestionsLoadErrorMessage, error);
        this._updateState({ isActiveSuggestionsLoading: false });
      });
  }

  private _applyCompletedFilter(nextFilter: IFilterState): void {
    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    this._fetchCompletedSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      filter: nextFilter
    })
      .then(() => {
        this._updateState({ isCompletedSuggestionsLoading: false });
      })
      .catch((error) => {
        this._handleError(strings.CompletedSuggestionsLoadErrorMessage, error);
        this._updateState({ isCompletedSuggestionsLoading: false });
      });
  }

  private _addSuggestion = async (): Promise<void> => {
    const title: string = this.state.newTitle.trim();
    const description: string = this.state.newDescription.trim();
    const category: SuggestionCategory = this.state.newCategory;
    const selectedSubcategory: ISubcategoryDefinition | undefined = this._getSelectedSubcategoryDefinition();

    if (!title) {
      this._handleError(strings.SuggestionTitleMissingMessage);
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const payload: IGraphSuggestionItemFields = {
        Title: title,
        Details: description,
        Status: this.state.defaultStatus,
        Category: category
      };

      if (selectedSubcategory) {
        payload.Subcategory = selectedSubcategory.title;
      }

      await this.props.graphService.addSuggestion(listId, payload);

      const defaultCategory: SuggestionCategory = this._getDefaultCategory(this.state.categories);
      this._debouncedSimilarSuggestionsSearch.cancel();

      this._updateState({
        newTitle: '',
        newDescription: '',
        newCategory: defaultCategory,
        newSubcategoryKey: this._getValidSubcategoryKeyForCategory(
          defaultCategory,
          undefined
        ),
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        isSimilarSuggestionsLoading: false,
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY }
      });

      await this._loadSuggestions();

      this._updateState({ success: 'Your suggestion has been added.' });
    } catch (error) {
      this._handleError(strings.SuggestionAddErrorMessage, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  };

  private async _toggleVote(item: ISuggestionItem): Promise<void> {
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    if (!normalizedUser) {
      this._handleError(strings.CurrentUserMissingErrorMessage);
      return;
    }

    const currentVote: IVoteEntry | undefined = item.voteEntries.find((vote) => vote.username === normalizedUser);
    const hasVoted: boolean = !!currentVote && currentVote.votes > 0;
    const remainingVotesForCategory: number = this._getRemainingVotesForCategory(item.category);

    if (!hasVoted && !this.state.isUnlimitedVotes && remainingVotesForCategory <= 0) {
      this._handleError(strings.NoVotesRemainingForCategoryMessage);
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const voteListId: string = this._getResolvedVotesListId();

      if (hasVoted && currentVote) {
        await this.props.graphService.deleteVote(voteListId, currentVote.id);
      } else {
        await this.props.graphService.addVote(voteListId, {
          SuggestionId: item.id,
          Username: normalizedUser,
          Votes: 1
        });
      }

      const [syncedVoteEntries] = await Promise.all([
        this._syncSuggestionVotes(item.id),
        this._refreshActiveSuggestions(),
        this._loadAvailableVotes()
      ]);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
      }

      if (syncedVoteEntries) {
        this._updateSimilarSuggestionsVotes(item.id, syncedVoteEntries);
      }
    } catch (error) {
      this._handleError(strings.VoteUpdateErrorMessage, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _syncSuggestionVotes(suggestionId: number): Promise<IVoteEntry[] | undefined> {
    if (!this._currentVotesListId) {
      return undefined;
    }

    try {
      const voteListId: string = this._getResolvedVotesListId();
      const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
        suggestionIds: [suggestionId]
      });
      const votesBySuggestion: Map<number, IVoteEntry[]> = this._groupVotesBySuggestion(voteItems);
      const voteEntries: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];
      const totalVotes: number = voteEntries.reduce((total, entry) => total + entry.votes, 0);
      const listId: string = this._getResolvedListId();

      await this.props.graphService.updateSuggestion(listId, suggestionId, {
        Votes: totalVotes
      });
      return voteEntries;
    } catch (error) {
      console.warn('Failed to sync suggestion votes.', error);
      return undefined;
    }
  }

  private _updateSimilarSuggestionsVotes(suggestionId: number, voteEntries: IVoteEntry[]): void {
    const voters: string[] = voteEntries.map((entry) => entry.username);
    const liveVotes: number = voteEntries.reduce((total, entry) => total + entry.votes, 0);

    this._updateState((prevState) => {
      const updatedItems: ISuggestionItem[] = prevState.similarSuggestions.items.map((item) => {
        if (item.id !== suggestionId) {
          return item;
        }

        const isCompleted: boolean = this._isCompletedStatusValue(
          item.status,
          prevState.completedStatus,
          prevState.deniedStatus
        );

        return {
          ...item,
          voteEntries,
          voters,
          votes: isCompleted ? item.votes : liveVotes
        };
      });

      const updatedSelected: ISuggestionItem | undefined =
        prevState.selectedSimilarSuggestion && prevState.selectedSimilarSuggestion.id === suggestionId
          ? {
              ...prevState.selectedSimilarSuggestion,
              voteEntries,
              voters,
              votes: this._isCompletedStatusValue(
                prevState.selectedSimilarSuggestion.status,
                prevState.completedStatus,
                prevState.deniedStatus
              )
                ? prevState.selectedSimilarSuggestion.votes
                : liveVotes
            }
          : prevState.selectedSimilarSuggestion;

      return {
        similarSuggestions: { ...prevState.similarSuggestions, items: updatedItems },
        selectedSimilarSuggestion: updatedSelected
      };
    });
  }

  private _canCurrentUserDeleteSuggestion(item: ISuggestionItem): boolean {
    const isCompleted: boolean = this._isCompletedStatusValue(
      item.status,
      this.state.completedStatus,
      this.state.deniedStatus
    );

    if (this.props.isCurrentUserAdmin) {
      return true;
    }

    if (isCompleted) {
      return false;
    }

    return this._isCurrentUserSuggestionOwner(item);
  }

  private _isCurrentUserSuggestionOwner(item: ISuggestionItem): boolean {
    const ownerLoginName: string | undefined = item.createdByLoginName;
    const currentUserLoginName: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    return !!ownerLoginName && !!currentUserLoginName && ownerLoginName === currentUserLoginName;
  }

  private _normalizeLoginName(value?: string): string | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const trimmed: string = value.trim();
    return trimmed.length > 0 ? trimmed.toLowerCase() : undefined;
  }

  private _isCommentSectionExpanded(suggestionId: number): boolean {
    return this.state.expandedCommentIds.indexOf(suggestionId) !== -1;
  }

  private _isCommentComposerVisible(suggestionId: number): boolean {
    return this.state.commentComposerIds.indexOf(suggestionId) !== -1;
  }

  private _isCommentSubmitting(suggestionId: number): boolean {
    return this.state.submittingCommentIds.indexOf(suggestionId) !== -1;
  }

  private _getCommentDraft(suggestionId: number): string {
    return this.state.commentDrafts[suggestionId] ?? '';
  }

  private _toggleCommentsSection = (suggestionId: number): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState(
      (prevState) => {
        const isExpanded: boolean = prevState.expandedCommentIds.indexOf(suggestionId) !== -1;
        const nextExpanded: number[] = isExpanded
          ? prevState.expandedCommentIds.filter((id) => id !== suggestionId)
          : [...prevState.expandedCommentIds, suggestionId];

        return { expandedCommentIds: nextExpanded };
      },
      () => {
        if (this._isCommentSectionExpanded(suggestionId)) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._ensureCommentsLoaded(suggestionId);
        }
      }
    );
  };

  private _toggleCommentComposer = (suggestionId: number): void => {
    this._updateState(
      (prevState) => {
        const isVisible: boolean = prevState.commentComposerIds.indexOf(suggestionId) !== -1;
        const nextComposerIds: number[] = isVisible
          ? prevState.commentComposerIds.filter((id) => id !== suggestionId)
          : [...prevState.commentComposerIds, suggestionId];

        return { commentComposerIds: nextComposerIds };
      },
      () => {
        if (this._isCommentComposerVisible(suggestionId)) {
          this._ensureCommentSectionExpanded(suggestionId);
        }
      }
    );
  };

  private _setCommentDraft(suggestionId: number, value: string): void {
    const nextValue: string = value ?? '';
    const hasExistingDraft: boolean = Object.prototype.hasOwnProperty.call(
      this.state.commentDrafts,
      suggestionId
    );

    if (nextValue.length === 0 && !hasExistingDraft) {
      return;
    }

    this._updateState((prevState) => {
      const drafts: Record<number, string> = { ...prevState.commentDrafts };

      if (nextValue.length === 0) {
        delete drafts[suggestionId];
      } else {
        drafts[suggestionId] = nextValue;
      }

      return { commentDrafts: drafts };
    });
  }

  private _omitCommentDraft(
    drafts: Record<number, string>,
    suggestionId: number
  ): Record<number, string> {
    if (!Object.prototype.hasOwnProperty.call(drafts, suggestionId)) {
      return drafts;
    }

    const nextDrafts: Record<number, string> = { ...drafts };
    delete nextDrafts[suggestionId];
    return nextDrafts;
  }

  private _handleCommentDraftChange = (item: ISuggestionItem, value: string): void => {
    this._setCommentDraft(item.id, value);
  };

  private _ensureCommentSectionExpanded(suggestionId: number): void {
    if (!this._isMounted) {
      return;
    }

    this.setState(
      (prevState) => {
        if (prevState.expandedCommentIds.indexOf(suggestionId) !== -1) {
          return null;
        }

        return {
          expandedCommentIds: [...prevState.expandedCommentIds, suggestionId]
        };
      },
      () => {
        if (this._isCommentSectionExpanded(suggestionId)) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._ensureCommentsLoaded(suggestionId);
        }
      }
    );
  }

  private async _ensureCommentsLoaded(suggestionId: number): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    const suggestion: ISuggestionItem | undefined = this._findSuggestionById(suggestionId);

    if (!suggestion) {
      return;
    }

    const hasLoadedComments: boolean = suggestion.areCommentsLoaded === true;

    if (hasLoadedComments) {
      return;
    }

    if (this.state.loadingCommentIds.indexOf(suggestionId) !== -1) {
      return;
    }

    this.setState((prevState) => ({
      loadingCommentIds: [...prevState.loadingCommentIds, suggestionId]
    }));

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      const commentItems: IGraphCommentItem[] = await this.props.graphService.getCommentItems(commentListId, {
        suggestionIds: [suggestionId]
      });
      const commentsBySuggestion: Map<number, ISuggestionComment[]> = this._groupCommentsBySuggestion(
        commentItems
      );
      const comments: ISuggestionComment[] = commentsBySuggestion.get(suggestionId) ?? [];

      this.setState((prevState) => ({
        loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== suggestionId),
        activeSuggestions: this._updateSuggestionItem(prevState.activeSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        completedSuggestions: this._updateSuggestionItem(prevState.completedSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        myVoteSuggestions: this._updateSuggestionArray(prevState.myVoteSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        adminSuggestions: this._updateSuggestionArray(prevState.adminSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        selectedSimilarSuggestion:
          prevState.selectedSimilarSuggestion && prevState.selectedSimilarSuggestion.id === suggestionId
            ? {
                ...prevState.selectedSimilarSuggestion,
                comments,
                commentCount: comments.length,
                areCommentsLoaded: true
              }
            : prevState.selectedSimilarSuggestion
      }));
    } catch (error) {
      this._handleError(strings.CommentsLoadErrorMessage, error);
      this.setState((prevState) => ({
        loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== suggestionId)
      }));
    }
  }

  private _findSuggestionById(suggestionId: number): ISuggestionItem | undefined {
    const {
      activeSuggestions,
      completedSuggestions,
      myVoteSuggestions,
      adminSuggestions,
      selectedSimilarSuggestion
    } = this.state;
    return (
      activeSuggestions.items.find((item) => item.id === suggestionId) ??
      completedSuggestions.items.find((item) => item.id === suggestionId) ??
      myVoteSuggestions.find((item) => item.id === suggestionId) ??
      adminSuggestions.find((item) => item.id === suggestionId) ??
      (selectedSimilarSuggestion && selectedSimilarSuggestion.id === suggestionId
        ? selectedSimilarSuggestion
        : undefined)
    );
  }

  private _updateSuggestionItem(
    source: IPaginatedSuggestionsState,
    suggestionId: number,
    updates: Partial<ISuggestionItem>
  ): IPaginatedSuggestionsState {
    const items: ISuggestionItem[] = source.items.map((item) =>
      item.id === suggestionId ? { ...item, ...updates } : item
    );

    return { ...source, items };
  }

  private _updateSuggestionArray(
    source: ISuggestionItem[],
    suggestionId: number,
    updates: Partial<ISuggestionItem>
  ): ISuggestionItem[] {
    return source.map((item) => (item.id === suggestionId ? { ...item, ...updates } : item));
  }

  private async _submitCommentForSuggestion(item: ISuggestionItem): Promise<void> {
    const draft: string = this._getCommentDraft(item.id);
    if (isRichTextValueEmpty(draft)) {
      this._handleError(strings.CommentMissingMessage);
      return;
    }

    this._updateState((prevState) => {
      const isAlreadySubmitting: boolean = prevState.submittingCommentIds.indexOf(item.id) !== -1;
      const submittingCommentIds: number[] = isAlreadySubmitting
        ? prevState.submittingCommentIds
        : [...prevState.submittingCommentIds, item.id];

      return { isLoading: true, error: undefined, success: undefined, submittingCommentIds };
    });

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      const title: string = `Suggestion #${item.id}`;

      await this.props.graphService.addCommentItem(commentListId, {
        Title: title.length > 255 ? title.slice(0, 255) : title,
        SuggestionId: item.id,
        Comment: draft
      });

      await this._refreshActiveSuggestions();
      this._ensureCommentSectionExpanded(item.id);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
        this._ensureCommentSectionExpanded(item.id);
      }

      this._updateState((prevState) => ({
        success: strings.CommentAddedMessage,
        commentDrafts: this._omitCommentDraft(prevState.commentDrafts, item.id),
        commentComposerIds: prevState.commentComposerIds.filter((id) => id !== item.id)
      }));
    } catch (error) {
      this._handleError(strings.CommentAddErrorMessage, error);
    } finally {
      this._updateState((prevState) => ({
        isLoading: false,
        submittingCommentIds: prevState.submittingCommentIds.filter((id) => id !== item.id)
      }));
    }
  }

  private async _deleteCommentFromSuggestion(
    item: ISuggestionItem,
    comment: ISuggestionComment
  ): Promise<void> {
    if (
      this._isCompletedStatusValue(
        item.status,
        this.state.completedStatus,
        this.state.deniedStatus
      )
    ) {
      this._handleError(strings.CommentDeleteCompletedSuggestionErrorMessage);
      return;
    }

    if (!this.props.isCurrentUserAdmin) {
      this._handleError(strings.CommentDeletePermissionErrorMessage);
      return;
    }

    const confirmed: boolean = window.confirm('Are you sure you want to delete this comment?');

    if (!confirmed) {
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      await this.props.graphService.deleteCommentItem(commentListId, comment.id);

      if (
        this._isCompletedStatusValue(
          item.status,
          this.state.completedStatus,
          this.state.deniedStatus
        )
      ) {
        await this._refreshCompletedSuggestions();
      } else {
        await this._refreshActiveSuggestions();
      }

      this._ensureCommentSectionExpanded(item.id);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
        this._ensureCommentSectionExpanded(item.id);
      }
      this._updateState({ success: 'The comment has been removed.' });
    } catch (error) {
      this._handleError(strings.CommentDeleteErrorMessage, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private _normalizeRequestedStatus(status: string): string | undefined {
    const trimmed: string = (status ?? '').trim();

    if (!trimmed) {
      return undefined;
    }

    const match: string | undefined = this.state.statuses.find((entry) =>
      this._areStatusesEqual(entry, trimmed)
    );

    return match ?? trimmed;
  }

  private async _updateSuggestionStatus(item: ISuggestionItem, requestedStatus: string): Promise<void> {
    if (!this.props.isCurrentUserAdmin) {
      this._handleError(strings.SuggestionStatusPermissionErrorMessage);
      return;
    }

    const targetStatus: string | undefined = this._normalizeRequestedStatus(requestedStatus);

    if (!targetStatus) {
      this._handleError(strings.SuggestionStatusSelectionErrorMessage);
      return;
    }

    if (this._areStatusesEqual(item.status, targetStatus)) {
      return;
    }

    const isCurrentlyCompleted: boolean = this._isCompletedStatusValue(
      item.status,
      this.state.completedStatus,
      this.state.deniedStatus
    );
    const willBeCompleted: boolean = this._isCompletedStatusValue(
      targetStatus,
      this.state.completedStatus,
      this.state.deniedStatus
    );

    let commentText: string | undefined;

    if (willBeCompleted && !isCurrentlyCompleted) {
      const commentInput: string | null = window.prompt(
        'Add a comment for this suggestion (optional). Leave blank to skip.',
        ''
      );

      if (commentInput === null) {
        return;
      }

      commentText = commentInput.trim();
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const voteListId: string = this._getResolvedVotesListId();
      const commentListId: string = this._getResolvedCommentsListId();

      const updatePayload: Partial<IGraphSuggestionItemFields> = {
        Status: targetStatus
      };

      if (willBeCompleted) {
        updatePayload.Votes = item.votes;
        updatePayload.CompletedDateTime = new Date().toISOString();
      } else if (isCurrentlyCompleted) {
        (updatePayload as Record<string, unknown>).CompletedDateTime = null;
      }

      await this.props.graphService.updateSuggestion(listId, item.id, updatePayload);

      if (willBeCompleted) {
        await this.props.graphService.deleteVotesForSuggestion(voteListId, item.id);
      }

      if (commentText && commentText.length > 0) {
        const title: string = `Suggestion #${item.id}`;
        await this.props.graphService.addCommentItem(commentListId, {
          Title: title.length > 255 ? title.slice(0, 255) : title,
          SuggestionId: item.id,
          Comment: commentText
        });
      }

      await Promise.all([this._loadSuggestions(), this._loadAvailableVotes()]);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, targetStatus);
      }

      this._updateState({ success: strings.SuggestionStatusUpdatedMessage });
    } catch (error) {
      this._handleError(strings.SuggestionStatusUpdateErrorMessage, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _deleteSuggestion(item: ISuggestionItem): Promise<void> {
    if (!this._canCurrentUserDeleteSuggestion(item)) {
      this._handleError(strings.SuggestionDeletePermissionErrorMessage);
      return;
    }

    const confirmation: boolean = window.confirm(strings.SuggestionDeleteConfirmationMessage);

    if (!confirmation) {
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const voteListId: string = this._getResolvedVotesListId();
      const commentListId: string = this._getResolvedCommentsListId();

      await this.props.graphService.deleteSuggestion(listId, item.id);
      await Promise.all([
        this.props.graphService.deleteVotesForSuggestion(voteListId, item.id),
        this.props.graphService.deleteCommentsForSuggestion(commentListId, item.id)
      ]);

      if (
        this._isCompletedStatusValue(
          item.status,
          this.state.completedStatus,
          this.state.deniedStatus
        )
      ) {
        await this._refreshCompletedSuggestions();
      } else {
        await Promise.all([this._refreshActiveSuggestions(), this._loadAvailableVotes()]);
      }

      if (this.state.selectedSimilarSuggestion?.id === item.id && this._isMounted) {
        this.setState((prevState) => ({
          selectedSimilarSuggestion: undefined,
          isSelectedSimilarSuggestionLoading: false,
          expandedCommentIds: prevState.expandedCommentIds.filter((id) => id !== item.id),
          loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== item.id)
        }));
      }

      this._updateState({ success: strings.SuggestionDeleteSuccessMessage });
    } catch (error) {
      this._handleError(strings.SuggestionDeleteErrorMessage, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private get _listTitle(): string {
    return this._normalizeListTitle(this.props.listTitle);
  }

  private _normalizeVoteListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.props.listTitle);
    if (trimmed.length > 0) {
      return trimmed;
    }

    if (normalizedListTitle === DEFAULT_SUGGESTIONS_LIST_TITLE) {
      return strings.DefaultVotesListTitle;
    }

    return `${normalizedListTitle}${DEFAULT_VOTES_LIST_SUFFIX}`;
  }

  private get _voteListTitle(): string {
    return this._normalizeVoteListTitle(this.props.voteListTitle, this.props.listTitle);
  }

  private _normalizeCommentListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.props.listTitle);
    if (trimmed.length > 0) {
      return trimmed;
    }

    if (normalizedListTitle === DEFAULT_SUGGESTIONS_LIST_TITLE) {
      return strings.DefaultCommentsListTitle;
    }

    return `${normalizedListTitle}${DEFAULT_COMMENTS_LIST_SUFFIX}`;
  }

  private get _commentListTitle(): string {
    return this._normalizeCommentListTitle(this.props.commentListTitle, this.props.listTitle);
  }

  private _normalizeOptionalListTitle(value?: string): string | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const trimmed: string = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  private get _subcategoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.props.subcategoryListTitle);
  }

  private get _categoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.props.categoryListTitle);
  }

  private get _statusListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.props.statusListTitle);
  }

  private _parseBooleanValue(value: unknown): boolean {
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
  }

  private _parsePageSizeOption(option?: IDropdownOption): number | undefined {
    if (!option) {
      return undefined;
    }

    const { key } = option;

    if (typeof key === 'number' && Number.isFinite(key) && key > 0) {
      return Math.floor(key);
    }

    if (typeof key === 'string') {
      const parsed: number = parseInt(key, 10);
      if (Number.isFinite(parsed) && parsed > 0) {
        return parsed;
      }
    }

    return undefined;
  }

  private _parseVotes(value: unknown): number {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value;
    }

    if (typeof value === 'string') {
      const parsed: number = parseInt(value, 10);
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }

    return 0;
  }

  private _tryNormalizeCategory(value: unknown): SuggestionCategory | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const normalized: string = value.trim();

    if (!normalized) {
      return undefined;
    }

    return this._findCategoryMatch(normalized, this.state.categories) ?? normalized;
  }

  private _normalizeCategory(value: unknown): SuggestionCategory {
    return this._tryNormalizeCategory(value) ?? this._getDefaultCategory(this.state.categories);
  }

  private _getResolvedListId(): string {
    if (!this._currentListId) {
      throw new Error('The suggestions list has not been initialized yet.');
    }

    return this._currentListId;
  }

  private _getResolvedVotesListId(): string {
    if (!this._currentVotesListId) {
      throw new Error('The votes list has not been initialized yet.');
    }

    return this._currentVotesListId;
  }

  private _getResolvedCommentsListId(): string {
    if (!this._currentCommentsListId) {
      throw new Error('The comments list has not been initialized yet.');
    }

    return this._currentCommentsListId;
  }

  private _getResolvedCategoryListId(): string {
    if (!this._currentCategoryListId) {
      throw new Error('The category list has not been initialized yet.');
    }

    return this._currentCategoryListId;
  }

  private _getResolvedSubcategoryListId(): string {
    if (!this._currentSubcategoryListId) {
      throw new Error('The subcategory list has not been initialized yet.');
    }

    return this._currentSubcategoryListId;
  }

  private _getResolvedStatusListId(): string {
    if (!this._currentStatusListId) {
      throw new Error('The status list has not been initialized yet.');
    }

    return this._currentStatusListId;
  }

  private _handleError(message: string, error?: unknown): void {
    console.error(message, error);
    this._updateState({ error: message, success: undefined });
  }

  private _onTabSelect = (nextTab: MainTabKey): void => {
    if (nextTab === this.state.selectedMainTab) {
      return;
    }

    this._updateState({ selectedMainTab: nextTab }, () => {
      const normalized: MainTabKey = nextTab;
      if (normalized === 'active') {
        this._applyActiveFilter(this.state.activeFilter);
        return;
      }

      if (normalized === 'completed') {
        this._applyCompletedFilter(this.state.completedFilter);
        return;
      }

      if (normalized === 'admin') {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this._loadAdminSuggestions();
        return;
      }

      if (normalized === 'myVotes') {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this._loadAvailableVotes();
      }
    });
  };

  private _flushPendingSimilarSuggestionsSearch(): void {
    if (!this._pendingSimilarSuggestionsQuery || !this._currentListId) {
      return;
    }

    const pendingQuery: ISimilarSuggestionsQuery = this._pendingSimilarSuggestionsQuery;
    this._pendingSimilarSuggestionsQuery = undefined;
    this._debouncedSimilarSuggestionsSearch(pendingQuery);
  }

  private _updateState(
    state:
      | Partial<ISamverkansportalenState>
      | ((prevState: ISamverkansportalenState) => Partial<ISamverkansportalenState>),
    callback?: () => void
  ): void {
    if (!this._isMounted) {
      return;
    }

    if (typeof state === 'function') {
      this.setState(
        (prevState) =>
          state(prevState) as Pick<ISamverkansportalenState, keyof ISamverkansportalenState>,
        callback
      );
      return;
    }

    this.setState(
      state as Pick<ISamverkansportalenState, keyof ISamverkansportalenState>,
      callback
    );
  }

  private _parseNumericId(value: unknown): number | undefined {
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
  }
}
