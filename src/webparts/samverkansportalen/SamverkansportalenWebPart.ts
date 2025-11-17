import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SamverkansportalenWebPartStrings';
import Samverkansportalen from './components/Samverkansportalen';
import {
  DEFAULT_SUGGESTIONS_HEADER_SUBTITLE,
  DEFAULT_SUGGESTIONS_HEADER_TITLE,
  DEFAULT_SUGGESTIONS_LIST_TITLE,
  DEFAULT_STATUS_DEFINITIONS,
  DEFAULT_TOTAL_VOTES_PER_USER,
  ISamverkansportalenProps
} from './components/ISamverkansportalenProps';
import GraphSuggestionsService, {
  DEFAULT_CATEGORY_LIST_TITLE,
  DEFAULT_COMMENT_LIST_TITLE,
  DEFAULT_STATUS_LIST_TITLE,
  DEFAULT_SUBCATEGORY_LIST_TITLE
} from './services/GraphSuggestionsService';

type ListCreationType = 'suggestions' | 'votes' | 'comments' | 'subcategories' | 'categories' | 'statuses';

interface IStatusDropdownOption extends IPropertyPaneDropdownOption {
  data?: {
    sortOrder?: number;
  };
}

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
  useTableLayout?: boolean;
  subcategoryListTitle?: string;
  categoryListTitle?: string;
  statusListTitle?: string;
  voteListTitle?: string;
  commentListTitle?: string;
  selectedSubcategoryKey?: string;
  newSubcategoryTitle?: string;
  selectedCategoryKey?: string;
  newCategoryTitle?: string;
  selectedStatusKey?: string;
  newStatusTitle?: string;
  headerTitle: string;
  headerSubtitle: string;
  statusDefinitions?: string;
  completedStatus?: string;
  defaultStatus?: string;
  totalVotesPerUser?: string;
}

export default class SamverkansportalenWebPart extends BaseClientSideWebPart<ISamverkansportalenWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [
    { key: DEFAULT_SUGGESTIONS_LIST_TITLE, text: DEFAULT_SUGGESTIONS_LIST_TITLE }
  ];
  private _isLoadingLists: boolean = false;
  private _pendingListCreation?: ListCreationType;
  private _listCreationMessage?: string;
  private _graphService?: GraphSuggestionsService;
  private _subcategoryOptions: IPropertyPaneDropdownOption[] = [];
  private _isLoadingSubcategories: boolean = false;
  private _isMutatingSubcategories: boolean = false;
  private _subcategoryStatusMessage?: string;
  private _resolvedSubcategoryListId?: string;
  private _resolvedSubcategoryListTitle?: string;
  private _categoryOptions: IPropertyPaneDropdownOption[] = [];
  private _isLoadingCategories: boolean = false;
  private _isMutatingCategories: boolean = false;
  private _categoryStatusMessage?: string;
  private _resolvedCategoryListId?: string;
  private _resolvedCategoryListTitle?: string;
  private _statusOptions: IStatusDropdownOption[] = [];
  private _isLoadingStatuses: boolean = false;
  private _isMutatingStatuses: boolean = false;
  private _statusStatusMessage?: string;
  private _resolvedStatusListId?: string;
  private _resolvedStatusListTitle?: string;

  public render(): void {
    const statuses: string[] = this._getStatusDefinitions();
    const completedStatus: string = this._getCompletedStatus(statuses);
    const defaultStatus: string = this._getDefaultStatus(statuses, completedStatus);

    const element: React.ReactElement<ISamverkansportalenProps> = React.createElement(
      Samverkansportalen,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userLoginName: this.context.pageContext.user.loginName,
        isCurrentUserAdmin: this._isCurrentUserSiteAdmin,
        graphService: this._getGraphService(),
        listTitle: this._selectedListTitle,
        voteListTitle: this._selectedVoteListTitle,
        commentListTitle: this._selectedCommentListTitle,
        useTableLayout: this.properties.useTableLayout,
        subcategoryListTitle: this._selectedSubcategoryListTitle,
        categoryListTitle: this._selectedCategoryListTitle,
        statusListTitle: this._selectedStatusListTitle,
        headerTitle: this._normalizeHeaderText(
          this.properties.headerTitle,
          DEFAULT_SUGGESTIONS_HEADER_TITLE
        ),
        headerSubtitle: this._normalizeHeaderText(
          this.properties.headerSubtitle,
          DEFAULT_SUGGESTIONS_HEADER_SUBTITLE
        ),
        statuses,
        completedStatus,
        defaultStatus,
        totalVotesPerUser: this._getTotalVotesPerUserSetting()
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private get _isCurrentUserSiteAdmin(): boolean {
    const legacyContext: unknown = this.context.pageContext.legacyPageContext;

    if (!legacyContext || typeof legacyContext !== 'object') {
      return false;
    }

    const isSiteAdmin: unknown = (legacyContext as { isSiteAdmin?: unknown }).isSiteAdmin;
    return isSiteAdmin === true;
  }

  protected onInit(): Promise<void> {
    this.properties.listTitle = this._normalizeListTitle(this.properties.listTitle);
    this.properties.voteListTitle = this._normalizeVoteListTitle(
      this.properties.voteListTitle,
      this.properties.listTitle
    );
    this.properties.commentListTitle = this._normalizeCommentListTitle(
      this.properties.commentListTitle,
      this.properties.listTitle
    );
    this.properties.subcategoryListTitle = this._normalizeOptionalListTitle(this.properties.subcategoryListTitle);
    this.properties.categoryListTitle = this._normalizeOptionalListTitle(this.properties.categoryListTitle);
    this.properties.statusListTitle = this._normalizeOptionalListTitle(this.properties.statusListTitle);
    this.properties.headerTitle = this._normalizeHeaderText(
      this.properties.headerTitle,
      DEFAULT_SUGGESTIONS_HEADER_TITLE
    );
    this.properties.headerSubtitle = this._normalizeHeaderText(
      this.properties.headerSubtitle,
      DEFAULT_SUGGESTIONS_HEADER_SUBTITLE
    );
    const normalizedStatusDefinitions: string = this._normalizeStatusDefinitions(
      this.properties.statusDefinitions
    );
    this.properties.statusDefinitions = normalizedStatusDefinitions;
    const statusList: string[] = this._parseStatusDefinitions(normalizedStatusDefinitions);
    const completedStatus: string = this._normalizeCompletedStatus(
      this.properties.completedStatus,
      statusList
    );
    this.properties.completedStatus = completedStatus;
    this.properties.defaultStatus = this._normalizeDefaultStatus(
      this.properties.defaultStatus,
      statusList,
      completedStatus
    );

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors,
      palette
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--accentColor', palette.themePrimary || null);
      this.domElement.style.setProperty('--accentColorTint', palette.themeLighter || null);
      this.domElement.style.setProperty('--accentColorLightest', palette.themeLighterAlt || null);
      this.domElement.style.setProperty('--accentColorDark', palette.themeDarkAlt || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationStart(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ensureListOptions();
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ensureSubcategoryOptions();
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ensureCategoryOptions();
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ensureStatusOptions();
    this._listCreationMessage = undefined;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._extendConfiguredLists();
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'commentListTitle') {
      this.properties.commentListTitle = this._normalizeCommentListTitle(
        typeof newValue === 'string' ? newValue : undefined,
        this.properties.listTitle
      );
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'subcategoryListTitle') {
      this.properties.subcategoryListTitle = this._normalizeOptionalListTitle(
        typeof newValue === 'string' ? newValue : undefined
      );
      this._resetSubcategoryState();
      this.context.propertyPane.refresh();

      this._ensureSubcategoryOptions().catch(() => {
        // Errors are handled inside _ensureSubcategoryOptions.
      });
    } else if (propertyPath === 'categoryListTitle') {
      this.properties.categoryListTitle = this._normalizeOptionalListTitle(
        typeof newValue === 'string' ? newValue : undefined
      );
      this._resetCategoryState();
      this.context.propertyPane.refresh();

      this._ensureCategoryOptions().catch(() => {
        // Errors are handled inside _ensureCategoryOptions.
      });
    } else if (propertyPath === 'statusListTitle') {
      this.properties.statusListTitle = this._normalizeOptionalListTitle(
        typeof newValue === 'string' ? newValue : undefined
      );
      this._resetStatusState();
      this.context.propertyPane.refresh();

      this._ensureStatusOptions().catch(() => {
        // Errors are handled inside _ensureStatusOptions.
      });
    } else if (propertyPath === 'headerTitle') {
      this.properties.headerTitle = this._normalizeHeaderText(
        typeof newValue === 'string' ? newValue : undefined,
        DEFAULT_SUGGESTIONS_HEADER_TITLE
      );
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'headerSubtitle') {
      this.properties.headerSubtitle = this._normalizeHeaderText(
        typeof newValue === 'string' ? newValue : undefined,
        DEFAULT_SUGGESTIONS_HEADER_SUBTITLE
      );
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'completedStatus') {
      const statuses: string[] =
        this._statusOptions.length > 0
          ? this._statusOptions
              .map((option) =>
                typeof option.text === 'string' && option.text.trim().length > 0
                  ? option.text.trim()
                  : option.key.toString()
              )
              .filter((status) => status.length > 0)
          : this._getStatusDefinitions();
      const completedStatus: string = this._normalizeCompletedStatus(
        typeof newValue === 'string' ? newValue : undefined,
        statuses
      );
      this.properties.completedStatus = completedStatus;
      this.properties.defaultStatus = this._normalizeDefaultStatus(
        this.properties.defaultStatus,
        statuses,
        completedStatus
      );
      this.context.propertyPane.refresh();
    } else if (propertyPath === 'defaultStatus') {
      const statuses: string[] =
        this._statusOptions.length > 0
          ? this._statusOptions
              .map((option) =>
                typeof option.text === 'string' && option.text.trim().length > 0
                  ? option.text.trim()
                  : option.key.toString()
              )
              .filter((status) => status.length > 0)
          : this._getStatusDefinitions();
      const completedStatus: string = this._normalizeCompletedStatus(
        this.properties.completedStatus,
        statuses
      );
      this.properties.completedStatus = completedStatus;
      this.properties.defaultStatus = this._normalizeDefaultStatus(
        typeof newValue === 'string' ? newValue : undefined,
        statuses,
        completedStatus
      );
      this.context.propertyPane.refresh();
    }
  }

  private async _ensureListOptions(): Promise<void> {
    if (this._isLoadingLists) {
      return;
    }

    this._isLoadingLists = true;

    try {
      const lists: string[] = (await this._getGraphService().getVisibleLists())
        .map((list) => list.displayName);

      const options: IPropertyPaneDropdownOption[] = lists
        .filter((title) => typeof title === 'string' && title.trim().length > 0)
        .map((title) => ({ key: title, text: title.trim() }))
        .sort((a, b) => a.text.localeCompare(b.text));

      const knownTitles: Set<string> = new Set(options.map((option) => option.key.toString()));
      const ensureOption = (title: string): void => {
        if (!knownTitles.has(title)) {
          knownTitles.add(title);
          options.push({ key: title, text: title });
        }
      };

      ensureOption(this._selectedListTitle);
      ensureOption(DEFAULT_SUGGESTIONS_LIST_TITLE);
      const selectedSubcategoryList: string | undefined = this._selectedSubcategoryListTitle;
      if (selectedSubcategoryList) {
        ensureOption(selectedSubcategoryList);
      }
      const selectedCategoryList: string | undefined = this._selectedCategoryListTitle;
      if (selectedCategoryList) {
        ensureOption(selectedCategoryList);
      }
      ensureOption(DEFAULT_CATEGORY_LIST_TITLE);
      ensureOption(this._selectedVoteListTitle);
      ensureOption(this._selectedCommentListTitle);
      ensureOption(DEFAULT_COMMENT_LIST_TITLE);

      options.sort((a, b) => a.text.localeCompare(b.text));

      this._listOptions = options;
    } catch (error) {
      console.error('Failed to load available lists for the property pane.', error);
    } finally {
      this._isLoadingLists = false;
      this.context.propertyPane.refresh();
    }
  }

  private _addListOption(title: string | undefined): void {
    const trimmed: string = (title ?? '').trim();

    if (!trimmed) {
      return;
    }

    if (this._listOptions.some((option) => option.key.toString() === trimmed)) {
      return;
    }

    this._listOptions = [...this._listOptions, { key: trimmed, text: trimmed }]
      .sort((a, b) => a.text.localeCompare(b.text));
  }

  private _resetSubcategoryState(): void {
    this._subcategoryOptions = [];
    this._subcategoryStatusMessage = undefined;
    this._resolvedSubcategoryListId = undefined;
    this._resolvedSubcategoryListTitle = undefined;
    this.properties.selectedSubcategoryKey = undefined;
    this.properties.newSubcategoryTitle = undefined;
  }

  private _resetCategoryState(): void {
    this._categoryOptions = [];
    this._categoryStatusMessage = undefined;
    this._resolvedCategoryListId = undefined;
    this._resolvedCategoryListTitle = undefined;
    this.properties.selectedCategoryKey = undefined;
    this.properties.newCategoryTitle = undefined;
  }

  private _resetStatusState(): void {
    this._statusOptions = [];
    this._statusStatusMessage = undefined;
    this._resolvedStatusListId = undefined;
    this._resolvedStatusListTitle = undefined;
    this.properties.selectedStatusKey = undefined;
    this.properties.newStatusTitle = undefined;
  }

  private async _ensureSubcategoryOptions(): Promise<void> {
    if (this._isLoadingSubcategories) {
      return;
    }

    const listTitle: string | undefined = this._selectedSubcategoryListTitle;

    if (!listTitle) {
      this._resetSubcategoryState();
      this._subcategoryStatusMessage = strings.SubcategoryListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isLoadingSubcategories = true;

    try {
      const listId: string | undefined = await this._getResolvedSubcategoryListId(listTitle);

      if (!listId) {
        this._subcategoryOptions = [];
        if (!this._subcategoryStatusMessage) {
          this._subcategoryStatusMessage = strings.SubcategoryLoadErrorMessage;
        }
        return;
      }

      const items = await this._getGraphService().getSubcategoryItems(listId);

      this._subcategoryOptions = items
        .map((item) => {
          const title: string = (item.fields?.Title ?? '').toString().trim();
          const category: string | undefined =
            typeof item.fields?.Category === 'string' && item.fields.Category.trim().length > 0
              ? item.fields.Category.trim()
              : undefined;

          if (!title) {
            return undefined;
          }

          const text: string = category ? `${title} (${category})` : title;

          return {
            key: item.id.toString(),
            text
          } as IPropertyPaneDropdownOption;
        })
        .filter((option): option is IPropertyPaneDropdownOption => !!option)
        .sort((a, b) => a.text.localeCompare(b.text));

      const availableKeys: Set<string> = new Set(
        this._subcategoryOptions.map((option) => option.key.toString())
      );

      const currentKey: string | undefined = this.properties.selectedSubcategoryKey;

      if (!currentKey || !availableKeys.has(currentKey)) {
        const firstKey: string | undefined = this._subcategoryOptions[0]?.key?.toString();
        this.properties.selectedSubcategoryKey = firstKey;
      }
    } catch (error) {
      console.error('Failed to load subcategories for the property pane.', error);
      this._subcategoryStatusMessage = strings.SubcategoryLoadErrorMessage;
    } finally {
      this._isLoadingSubcategories = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _ensureCategoryOptions(): Promise<void> {
    if (this._isLoadingCategories) {
      return;
    }

    const listTitle: string | undefined = this._selectedCategoryListTitle;

    if (!listTitle) {
      this._resetCategoryState();
      this._categoryStatusMessage = strings.CategoryListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isLoadingCategories = true;

    try {
      const listId: string | undefined = await this._getResolvedCategoryListId(listTitle);

      if (!listId) {
        this._categoryOptions = [];
        if (!this._categoryStatusMessage) {
          this._categoryStatusMessage = strings.CategoryLoadErrorMessage;
        }
        return;
      }

      const items = await this._getGraphService().getCategoryItems(listId);

      this._categoryOptions = items
        .map((item) => {
          const title: string = (item.fields?.Title ?? '').toString().trim();

          if (!title) {
            return undefined;
          }

          return {
            key: item.id.toString(),
            text: title
          } as IPropertyPaneDropdownOption;
        })
        .filter((option): option is IPropertyPaneDropdownOption => !!option)
        .sort((a, b) => a.text.localeCompare(b.text));

      const availableKeys: Set<string> = new Set(this._categoryOptions.map((option) => option.key.toString()));

      const currentKey: string | undefined = this.properties.selectedCategoryKey;

      if (!currentKey || !availableKeys.has(currentKey)) {
        const firstKey: string | undefined = this._categoryOptions[0]?.key?.toString();
        this.properties.selectedCategoryKey = firstKey;
      }
    } catch (error) {
      console.error('Failed to load categories for the property pane.', error);
      this._categoryStatusMessage = strings.CategoryLoadErrorMessage;
    } finally {
      this._isLoadingCategories = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _ensureStatusOptions(): Promise<void> {
    if (this._isLoadingStatuses) {
      return;
    }

    const listTitle: string | undefined = this._selectedStatusListTitle;

    if (!listTitle) {
      this._resetStatusState();
      this._statusStatusMessage = strings.StatusListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isLoadingStatuses = true;

    try {
      const listId: string | undefined = await this._getResolvedStatusListId(listTitle);

      if (!listId) {
        this._statusOptions = [];
        this.properties.selectedStatusKey = undefined;
        if (!this._statusStatusMessage) {
          this._statusStatusMessage = strings.StatusLoadErrorMessage;
        }
        return;
      }

      const items = await this._getGraphService().getStatusItems(listId);

      const definitions: Array<{ id: number; title: string; order: number | undefined }> = [];

      items.forEach((item) => {
        const id: number = item.id;
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

        definitions.push({ id, title, order });
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

      this._statusOptions = definitions.map(({ id, title, order }) => ({
        key: id.toString(),
        text: title,
        data: { sortOrder: order }
      }));

      if (this._statusOptions.length > 0) {
        const availableKeys: Set<string> = new Set(
          this._statusOptions.map((option) => option.key.toString())
        );
        const currentKey: string | undefined = this.properties.selectedStatusKey;

        if (!currentKey || !availableKeys.has(currentKey)) {
          const firstKey: string | undefined = this._statusOptions[0]?.key?.toString();
          this.properties.selectedStatusKey = firstKey;
        }
      } else {
        this.properties.selectedStatusKey = undefined;
      }

      const availableStatuses: string[] = this._statusOptions.map((option) =>
        typeof option.text === 'string' && option.text.trim().length > 0
          ? option.text
          : option.key.toString()
      );
      this.properties.completedStatus = this._normalizeCompletedStatus(
        this.properties.completedStatus,
        availableStatuses
      );
      this._statusStatusMessage = undefined;
    } catch (error) {
      console.error('Failed to load statuses for the property pane.', error);
      this._statusStatusMessage = strings.StatusLoadErrorMessage;
    } finally {
      this._isLoadingStatuses = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _getResolvedSubcategoryListId(listTitle?: string): Promise<string | undefined> {
    const normalizedTitle: string | undefined = this._normalizeOptionalListTitle(
      listTitle ?? this._selectedSubcategoryListTitle
    );

    if (!normalizedTitle) {
      return undefined;
    }

    if (
      this._resolvedSubcategoryListId &&
      this._resolvedSubcategoryListTitle &&
      this._resolvedSubcategoryListTitle.localeCompare(normalizedTitle, undefined, {
        sensitivity: 'accent'
      }) === 0
    ) {
      return this._resolvedSubcategoryListId;
    }

    const listInfo = await this._getGraphService().getListByTitle(normalizedTitle);

    if (!listInfo) {
      return undefined;
    }

    this._resolvedSubcategoryListId = listInfo.id;
    this._resolvedSubcategoryListTitle = normalizedTitle;
    return listInfo.id;
  }

  private async _getResolvedCategoryListId(listTitle?: string): Promise<string | undefined> {
    const normalizedTitle: string | undefined = this._normalizeOptionalListTitle(
      listTitle ?? this._selectedCategoryListTitle
    );

    if (!normalizedTitle) {
      return undefined;
    }

    if (
      this._resolvedCategoryListId &&
      this._resolvedCategoryListTitle &&
      this._resolvedCategoryListTitle.localeCompare(normalizedTitle, undefined, {
        sensitivity: 'accent'
      }) === 0
    ) {
      return this._resolvedCategoryListId;
    }

    const listInfo = await this._getGraphService().getListByTitle(normalizedTitle);

    if (!listInfo) {
      return undefined;
    }

    this._resolvedCategoryListId = listInfo.id;
    this._resolvedCategoryListTitle = normalizedTitle;
    return listInfo.id;
  }

  private async _getResolvedStatusListId(listTitle?: string): Promise<string | undefined> {
    const normalizedTitle: string | undefined = this._normalizeOptionalListTitle(
      listTitle ?? this._selectedStatusListTitle
    );

    if (!normalizedTitle) {
      return undefined;
    }

    if (
      this._resolvedStatusListId &&
      this._resolvedStatusListTitle &&
      this._resolvedStatusListTitle.localeCompare(normalizedTitle, undefined, {
        sensitivity: 'accent'
      }) === 0
    ) {
      return this._resolvedStatusListId;
    }

    const listInfo = await this._getGraphService().getListByTitle(normalizedTitle);

    if (!listInfo) {
      return undefined;
    }

    this._resolvedStatusListId = listInfo.id;
    this._resolvedStatusListTitle = normalizedTitle;
    return listInfo.id;
  }

  private async _mutateSubcategoryList(
    executor: (listId: string) => Promise<string | undefined>
  ): Promise<void> {
    if (this._isMutatingSubcategories) {
      return;
    }

    const listId: string | undefined = await this._getResolvedSubcategoryListId();

    if (!listId) {
      this._subcategoryStatusMessage = strings.SubcategoryListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isMutatingSubcategories = true;
    this._subcategoryStatusMessage = strings.SubcategoryUpdateProgressMessage;
    this.context.propertyPane.refresh();

    try {
      const message: string | undefined = await executor(listId);
      const previousStatus: string | undefined = this._subcategoryStatusMessage;
      await this._ensureSubcategoryOptions();
      if (this._subcategoryStatusMessage === previousStatus) {
        this._subcategoryStatusMessage = message ?? strings.SubcategoryUpdateSuccessMessage;
      }
    } catch (error) {
      console.error('Failed to update the subcategory list.', error);
      this._subcategoryStatusMessage = strings.SubcategoryUpdateErrorMessage;
    } finally {
      this._isMutatingSubcategories = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _mutateCategoryList(
    executor: (listId: string) => Promise<string | undefined>
  ): Promise<void> {
    if (this._isMutatingCategories) {
      return;
    }

    const listId: string | undefined = await this._getResolvedCategoryListId();

    if (!listId) {
      this._categoryStatusMessage = strings.CategoryListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isMutatingCategories = true;
    this._categoryStatusMessage = strings.CategoryUpdateProgressMessage;
    this.context.propertyPane.refresh();

    try {
      const message: string | undefined = await executor(listId);
      const previousStatus: string | undefined = this._categoryStatusMessage;
      await this._ensureCategoryOptions();
      if (this._categoryStatusMessage === previousStatus) {
        this._categoryStatusMessage = message ?? strings.CategoryUpdateSuccessMessage;
      }
    } catch (error) {
      console.error('Failed to update the category list.', error);
      this._categoryStatusMessage = strings.CategoryUpdateErrorMessage;
    } finally {
      this._isMutatingCategories = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _mutateStatusList(
    executor: (listId: string) => Promise<string | undefined>
  ): Promise<void> {
    if (this._isMutatingStatuses) {
      return;
    }

    const listId: string | undefined = await this._getResolvedStatusListId();

    if (!listId) {
      this._statusStatusMessage = strings.StatusListNotConfiguredMessage;
      this.context.propertyPane.refresh();
      return;
    }

    this._isMutatingStatuses = true;
    this._statusStatusMessage = strings.StatusUpdateProgressMessage;
    this.context.propertyPane.refresh();

    try {
      const message: string | undefined = await executor(listId);
      const previousStatus: string | undefined = this._statusStatusMessage;
      await this._ensureStatusOptions();
      if (this._statusStatusMessage === previousStatus) {
        this._statusStatusMessage = message ?? strings.StatusUpdateSuccessMessage;
      }
    } catch (error) {
      console.error('Failed to update the status list.', error);
      this._statusStatusMessage = strings.StatusUpdateErrorMessage;
    } finally {
      this._isMutatingStatuses = false;
      this.context.propertyPane.refresh();
    }
  }

  private _getNextStatusSortOrder(): number | undefined {
    const orders: number[] = this._statusOptions
      .map((option) => option.data?.sortOrder)
      .filter((order): order is number => typeof order === 'number' && Number.isFinite(order));

    if (orders.length > 0) {
      return Math.max(...orders) + 1;
    }

    if (this._statusOptions.length > 0) {
      return this._statusOptions.length;
    }

    return 0;
  }

  private _handleAddSubcategoryClick = (): void => {
    this._addSubcategory().catch(() => {
      // Errors are handled in _mutateSubcategoryList.
    });
  };

  private async _addSubcategory(): Promise<void> {
    const title: string = (this.properties.newSubcategoryTitle ?? '').trim();

    if (!title) {
      this._subcategoryStatusMessage = strings.SubcategoryNameMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateSubcategoryList(async (listId) => {
      await this._getGraphService().addSubcategoryItem(listId, { Title: title });
      this.properties.newSubcategoryTitle = undefined;
      return strings.SubcategoryAddedMessage.replace('{0}', title);
    });
  }

  private _handleAddCategoryClick = (): void => {
    this._addCategory().catch(() => {
      // Errors are handled in _mutateCategoryList.
    });
  };

  private async _addCategory(): Promise<void> {
    const title: string = (this.properties.newCategoryTitle ?? '').trim();

    if (!title) {
      this._categoryStatusMessage = strings.CategoryNameMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateCategoryList(async (listId) => {
      await this._getGraphService().addCategoryItem(listId, { Title: title });
      this.properties.newCategoryTitle = undefined;
      return strings.CategoryAddedMessage.replace('{0}', title);
    });
  }

  private _handleAddStatusClick = (): void => {
    this._addStatus().catch(() => {
      // Errors are handled in _mutateStatusList.
    });
  };

  private async _addStatus(): Promise<void> {
    const title: string = (this.properties.newStatusTitle ?? '').trim();

    if (!title) {
      this._statusStatusMessage = strings.StatusNameMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateStatusList(async (listId) => {
      const nextOrder: number | undefined = this._getNextStatusSortOrder();
      const fields: { Title: string; SortOrder?: number } = { Title: title };

      if (typeof nextOrder === 'number' && Number.isFinite(nextOrder)) {
        fields.SortOrder = nextOrder;
      }

      await this._getGraphService().addStatusItem(listId, fields);
      this.properties.newStatusTitle = undefined;
      return strings.StatusAddedMessage.replace('{0}', title);
    });
  }

  private _handleRemoveCategoryClick = (): void => {
    this._removeCategory().catch(() => {
      // Errors are handled in _mutateCategoryList.
    });
  };

  private async _removeCategory(): Promise<void> {
    const key: string | undefined = this.properties.selectedCategoryKey;

    if (!key) {
      this._categoryStatusMessage = strings.CategorySelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    const parsedId: number = parseInt(key, 10);

    if (!Number.isFinite(parsedId)) {
      this._categoryStatusMessage = strings.CategorySelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateCategoryList(async (listId) => {
      await this._getGraphService().deleteCategoryItem(listId, parsedId);
      return strings.CategoryRemovedMessage;
    });
  }

  private _handleRemoveStatusClick = (): void => {
    this._removeStatus().catch(() => {
      // Errors are handled in _mutateStatusList.
    });
  };

  private async _removeStatus(): Promise<void> {
    const key: string | undefined = this.properties.selectedStatusKey;

    if (!key) {
      this._statusStatusMessage = strings.StatusSelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    const parsedId: number = parseInt(key, 10);

    if (!Number.isFinite(parsedId)) {
      this._statusStatusMessage = strings.StatusSelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateStatusList(async (listId) => {
      await this._getGraphService().deleteStatusItem(listId, parsedId);
      return strings.StatusRemovedMessage;
    });
  }

  private _handleRemoveSubcategoryClick = (): void => {
    this._removeSubcategory().catch(() => {
      // Errors are handled in _mutateSubcategoryList.
    });
  };

  private async _removeSubcategory(): Promise<void> {
    const key: string | undefined = this.properties.selectedSubcategoryKey;

    if (!key) {
      this._subcategoryStatusMessage = strings.SubcategorySelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    const parsedId: number = parseInt(key, 10);

    if (!Number.isFinite(parsedId)) {
      this._subcategoryStatusMessage = strings.SubcategorySelectionMissingMessage;
      this.context.propertyPane.refresh();
      return;
    }

    await this._mutateSubcategoryList(async (listId) => {
      await this._getGraphService().deleteSubcategoryItem(listId, parsedId);
      return strings.SubcategoryRemovedMessage;
    });
  }

  private _handleEnsureSuggestionsListClick = (): void => {
    this._ensureListFromPropertyPane('suggestions').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private _handleEnsureVoteListClick = (): void => {
    this._ensureListFromPropertyPane('votes').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private _handleEnsureCommentListClick = (): void => {
    this._ensureListFromPropertyPane('comments').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private _handleEnsureSubcategoryListClick = (): void => {
    this._ensureListFromPropertyPane('subcategories').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private _handleEnsureCategoryListClick = (): void => {
    this._ensureListFromPropertyPane('categories').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private _handleEnsureStatusListClick = (): void => {
    this._ensureListFromPropertyPane('statuses').catch(() => {
      // Errors are handled inside _ensureListFromPropertyPane.
    });
  };

  private async _ensureListFromPropertyPane(type: ListCreationType): Promise<void> {
    const promptLabel: string = this._getListPromptLabel(type);
    const promptMessage: string = strings.CreateListPromptMessage.replace('{0}', promptLabel);
    const defaultName: string = this._getDefaultListName(type);
    const input: string | null = window.prompt(promptMessage, defaultName);

    if (input === null) {
      return;
    }

    const trimmed: string = input.trim();

    if (!trimmed) {
      this._setListCreationMessage(strings.CreateListNameMissingMessage);
      return;
    }

    this._pendingListCreation = type;
    this._setListCreationMessage(this._getListProgressMessage(type));

    let message: string | undefined;

    try {
      const service: GraphSuggestionsService = this._getGraphService();

      if (type === 'suggestions') {
        const result: { created: boolean } = await service.ensureList(trimmed);
        this.properties.listTitle = trimmed;
        this._addListOption(trimmed);

        if (result.created) {
          const defaultVoteListTitle: string = this._getDefaultVoteListTitle(trimmed);
          this.properties.voteListTitle = defaultVoteListTitle;
          this._addListOption(defaultVoteListTitle);
          const defaultCommentListTitle: string = this._getDefaultCommentListTitle(trimmed);
          this.properties.commentListTitle = defaultCommentListTitle;
          this._addListOption(defaultCommentListTitle);
        }

        this.render();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      } else if (type === 'votes') {
        const result: { created: boolean } = await service.ensureVoteList(trimmed);
        this.properties.voteListTitle = trimmed;
        this._addListOption(trimmed);
        this.render();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      } else if (type === 'comments') {
        const result: { created: boolean } = await service.ensureCommentList(trimmed);
        this.properties.commentListTitle = trimmed;
        this._addListOption(trimmed);
        this.render();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      } else if (type === 'subcategories') {
        const result: { id: string; created: boolean } = await service.ensureSubcategoryList(trimmed);
        this.properties.subcategoryListTitle = trimmed;
        this._addListOption(trimmed);
        this._resetSubcategoryState();
        this._resolvedSubcategoryListId = result.id;
        this._resolvedSubcategoryListTitle = trimmed;
        this.render();
        await this._ensureSubcategoryOptions();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      } else if (type === 'statuses') {
        const result: { id: string; created: boolean } = await service.ensureStatusList(trimmed);
        this.properties.statusListTitle = trimmed;
        this._addListOption(trimmed);
        this._resetStatusState();
        this._resolvedStatusListId = result.id;
        this._resolvedStatusListTitle = trimmed;
        this.render();
        await this._ensureStatusOptions();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      } else {
        const result: { id: string; created: boolean } = await service.ensureCategoryList(trimmed);
        this.properties.categoryListTitle = trimmed;
        this._addListOption(trimmed);
        this._resetCategoryState();
        this._resolvedCategoryListId = result.id;
        this._resolvedCategoryListTitle = trimmed;
        this.render();
        await this._ensureCategoryOptions();
        message = result.created
          ? strings.CreateListSuccessMessage.replace('{0}', trimmed)
          : strings.CreateListAlreadyExistsMessage;
      }

      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Failed to create or update the SharePoint list from the property pane.', error);
      message = strings.CreateListErrorMessage;
    } finally {
      this._pendingListCreation = undefined;
      this._setListCreationMessage(message);
    }
  }

  private _getListPromptLabel(type: ListCreationType): string {
    switch (type) {
      case 'votes':
        return strings.CreateListPromptVotesLabel;
      case 'comments':
        return strings.CreateListPromptCommentsLabel;
      case 'subcategories':
        return strings.CreateListPromptSubcategoryLabel;
      case 'categories':
        return strings.CreateListPromptCategoryLabel;
      case 'statuses':
        return strings.CreateListPromptStatusLabel;
      default:
        return strings.CreateListPromptSuggestionsLabel;
    }
  }

  private _getDefaultListName(type: ListCreationType): string {
    switch (type) {
      case 'votes':
        return this._selectedVoteListTitle;
      case 'comments':
        return this._selectedCommentListTitle;
      case 'subcategories':
        return this._selectedSubcategoryListTitle ?? DEFAULT_SUBCATEGORY_LIST_TITLE;
      case 'statuses':
        return this._selectedStatusListTitle ?? DEFAULT_STATUS_LIST_TITLE;
      case 'categories':
        return this._selectedCategoryListTitle ?? DEFAULT_CATEGORY_LIST_TITLE;
      default:
        return this._selectedListTitle;
    }
  }

  private _getListProgressMessage(type: ListCreationType): string {
    switch (type) {
      case 'votes':
        return strings.CreateVotesListProgressMessage;
      case 'comments':
        return strings.CreateCommentsListProgressMessage;
      case 'subcategories':
        return strings.CreateSubcategoryListProgressMessage;
      case 'statuses':
        return strings.CreateStatusListProgressMessage;
      case 'categories':
        return strings.CreateCategoryListProgressMessage;
      default:
        return strings.CreateSuggestionsListProgressMessage;
    }
  }

  private _setListCreationMessage(message?: string): void {
    this._listCreationMessage = message;
    this.context.propertyPane.refresh();
  }

  private async _extendConfiguredLists(): Promise<void> {
    const listTitle: string = this._selectedListTitle;
    const voteListTitle: string = this._selectedVoteListTitle;
    const commentListTitle: string = this._selectedCommentListTitle;
    const subcategoryListTitle: string | undefined = this._selectedSubcategoryListTitle;
    const categoryListTitle: string | undefined = this._selectedCategoryListTitle;
    const statusListTitle: string | undefined = this._selectedStatusListTitle;

    try {
      await this._getGraphService().ensureList(listTitle);
      await this._getGraphService().ensureVoteList(voteListTitle);
      await this._getGraphService().ensureCommentList(commentListTitle);
      if (subcategoryListTitle) {
        await this._getGraphService().ensureSubcategoryList(subcategoryListTitle);
      }
      if (categoryListTitle) {
        await this._getGraphService().ensureCategoryList(categoryListTitle);
      }
      if (statusListTitle) {
        await this._getGraphService().ensureStatusList(statusListTitle);
      }
    } catch (error) {
      console.error('Failed to ensure the configured suggestions list.', error);
    }
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private get _selectedListTitle(): string {
    return this._normalizeListTitle(this.properties.listTitle);
  }

  private _getDefaultVoteListTitle(listTitle: string): string {
    const trimmed: string = listTitle.trim();
    return `${trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE}Votes`;
  }

  private _normalizeVoteListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.properties.listTitle);
    return trimmed.length > 0 ? trimmed : this._getDefaultVoteListTitle(normalizedListTitle);
  }

  private get _selectedVoteListTitle(): string {
    return this._normalizeVoteListTitle(this.properties.voteListTitle, this.properties.listTitle);
  }

  private _getDefaultCommentListTitle(listTitle: string): string {
    const trimmed: string = listTitle.trim();
    return `${trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE}Comments`;
  }

  private _normalizeCommentListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.properties.listTitle);
    return trimmed.length > 0 ? trimmed : this._getDefaultCommentListTitle(normalizedListTitle);
  }

  private get _selectedCommentListTitle(): string {
    return this._normalizeCommentListTitle(this.properties.commentListTitle, this.properties.listTitle);
  }

  private _normalizeOptionalListTitle(value?: string): string | undefined {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  private _normalizeHeaderText(value: string | undefined, fallback: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : fallback;
  }

  private _parseStatusDefinitions(value?: string): string[] {
    const source: string = typeof value === 'string' && value.trim().length > 0
      ? value
      : DEFAULT_STATUS_DEFINITIONS;
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
      return this._parseStatusDefinitions(DEFAULT_STATUS_DEFINITIONS);
    }

    return results.length > 0 ? results : ['Active', 'Done'];
  }

  private _normalizeStatusDefinitions(value?: string): string {
    const statuses: string[] = this._parseStatusDefinitions(value);
    return statuses.join('\n');
  }

  private _normalizeCompletedStatus(value: string | undefined, statuses: string[]): string {
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
  }

  private _normalizeDefaultStatus(
    value: string | undefined,
    statuses: string[],
    completedStatus: string
  ): string {
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
  }

  private _normalizeTotalVotesPerUser(value: string | undefined): number {
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
  }

  private _getTotalVotesPerUserSetting(): number {
    return this._normalizeTotalVotesPerUser(this.properties.totalVotesPerUser);
  }

  private _validateTotalVotesPerUser = (value: string): string => {
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

  private _getDropdownOptionText(option: IPropertyPaneDropdownOption): string {
    const text: string | undefined =
      typeof option.text === 'string' ? option.text.trim() : undefined;

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
  }

  private _getStatusDefinitions(): string[] {
    return this._parseStatusDefinitions(this.properties.statusDefinitions);
  }

  private _getCompletedStatus(statuses: string[]): string {
    return this._normalizeCompletedStatus(this.properties.completedStatus, statuses);
  }

  private _getDefaultStatus(statuses: string[], completedStatus: string): string {
    return this._normalizeDefaultStatus(this.properties.defaultStatus, statuses, completedStatus);
  }

  private get _selectedSubcategoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.properties.subcategoryListTitle);
  }

  private get _selectedCategoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.properties.categoryListTitle);
  }

  private get _selectedStatusListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.properties.statusListTitle);
  }

  private _getGraphService(): GraphSuggestionsService {
    if (!this._graphService) {
      this._graphService = new GraphSuggestionsService(
        this.context.msGraphClientFactory,
        this.context.pageContext.web.absoluteUrl
      );
    }

    return this._graphService;
  }

  private get _isListCreationInProgress(): boolean {
    return typeof this._pendingListCreation !== 'undefined';
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const hasSubcategoryListConfigured: boolean = !!this._selectedSubcategoryListTitle;
    const hasCategoryListConfigured: boolean = !!this._selectedCategoryListTitle;
    const subcategoryDropdownOptions: IPropertyPaneDropdownOption[] =
      this._subcategoryOptions.length > 0
        ? this._subcategoryOptions
        : [{ key: '__no_subcategories__', text: strings.SubcategoryDropdownPlaceholder }];
    const categoryDropdownOptions: IPropertyPaneDropdownOption[] =
      this._categoryOptions.length > 0
        ? this._categoryOptions
        : [{ key: '__no_categories__', text: strings.CategoryDropdownPlaceholder }];
    const canMutateSubcategories: boolean =
      hasSubcategoryListConfigured && !this._isLoadingSubcategories && !this._isMutatingSubcategories;
    const canMutateCategories: boolean =
      hasCategoryListConfigured && !this._isLoadingCategories && !this._isMutatingCategories;
    const canAddSubcategory: boolean =
      canMutateSubcategories && (this.properties.newSubcategoryTitle ?? '').trim().length > 0;
    const canRemoveSubcategory: boolean =
      canMutateSubcategories && !!this.properties.selectedSubcategoryKey && this._subcategoryOptions.length > 0;
    const canAddCategory: boolean =
      canMutateCategories && (this.properties.newCategoryTitle ?? '').trim().length > 0;
    const canRemoveCategory: boolean =
      canMutateCategories && !!this.properties.selectedCategoryKey && this._categoryOptions.length > 0;
    const hasStatusListConfigured: boolean = !!this._selectedStatusListTitle;
    const statusDropdownOptions: IPropertyPaneDropdownOption[] =
      this._statusOptions.length > 0
        ? this._statusOptions
        : [{ key: '__no_statuses__', text: strings.StatusDropdownPlaceholder }];
    const canMutateStatuses: boolean =
      hasStatusListConfigured && !this._isLoadingStatuses && !this._isMutatingStatuses;
    const canAddStatus: boolean =
      canMutateStatuses && (this.properties.newStatusTitle ?? '').trim().length > 0;
    const canRemoveStatus: boolean =
      canMutateStatuses && !!this.properties.selectedStatusKey && this._statusOptions.length > 0;
    const fallbackStatusDefinitions: string[] = this._getStatusDefinitions();
    const effectiveStatusOptions: IPropertyPaneDropdownOption[] =
      this._statusOptions.length > 0
        ? this._statusOptions.map((option) => {
            const text: string =
              typeof option.text === 'string' && option.text.trim().length > 0
                ? option.text
                : option.key.toString();
            return { key: text, text };
          })
        : fallbackStatusDefinitions.map((status) => ({ key: status, text: status }));
    const effectiveStatuses: string[] = effectiveStatusOptions.map((option) =>
      typeof option.text === 'string' && option.text.trim().length > 0
        ? option.text
        : option.key.toString()
    );
    const completedStatus: string = this._normalizeCompletedStatus(
      this.properties.completedStatus,
      effectiveStatuses
    );
    this.properties.completedStatus = completedStatus;
    const defaultStatus: string = this._normalizeDefaultStatus(
      this.properties.defaultStatus,
      effectiveStatuses,
      completedStatus
    );
    this.properties.defaultStatus = defaultStatus;
    const completedStatusOptions: IPropertyPaneDropdownOption[] =
      effectiveStatusOptions.length > 0
        ? effectiveStatusOptions
        : [{ key: completedStatus, text: completedStatus }];
    const defaultStatusOptions: IPropertyPaneDropdownOption[] = [];
    const defaultStatusKeys: Set<string> = new Set();
    const addDefaultStatusOption = (status: string | undefined): void => {
      const normalized: string = (status ?? '').trim();

      if (!normalized) {
        return;
      }

      const key: string = normalized.toLowerCase();

      if (defaultStatusKeys.has(key)) {
        return;
      }

      defaultStatusKeys.add(key);
      defaultStatusOptions.push({ key: normalized, text: normalized });
    };

    effectiveStatusOptions.forEach((option) => {
      addDefaultStatusOption(this._getDropdownOptionText(option));
    });
    addDefaultStatusOption(defaultStatus);

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('listTitle', {
                  label: strings.ListFieldLabel,
                  options: this._listOptions,
                  selectedKey: this._selectedListTitle,
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createSuggestionsListButton', {
                  text: strings.CreateSuggestionsListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureSuggestionsListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneDropdown('voteListTitle', {
                  label: strings.VoteListFieldLabel,
                  options: this._listOptions,
                  selectedKey: this._selectedVoteListTitle,
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createVoteListButton', {
                  text: strings.CreateVotesListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureVoteListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneTextField('totalVotesPerUser', {
                  label: strings.TotalVotesPerUserFieldLabel,
                  description: strings.TotalVotesPerUserFieldDescription.replace(
                    '{0}',
                    DEFAULT_TOTAL_VOTES_PER_USER.toString()
                  ),
                  value: this.properties.totalVotesPerUser ?? '',
                  placeholder: DEFAULT_TOTAL_VOTES_PER_USER.toString(),
                  validateOnFocusOut: true,
                  onGetErrorMessage: this._validateTotalVotesPerUser
                }),
                PropertyPaneDropdown('commentListTitle', {
                  label: strings.CommentListFieldLabel,
                  options: this._listOptions,
                  selectedKey: this._selectedCommentListTitle,
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createCommentListButton', {
                  text: strings.CreateCommentsListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureCommentListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneDropdown('categoryListTitle', {
                  label: strings.CategoryListFieldLabel,
                  options: [
                    { key: '', text: strings.CategoryListDefaultOptionLabel },
                    ...this._listOptions
                  ],
                  selectedKey: this._selectedCategoryListTitle ?? '',
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createCategoryListButton', {
                  text: strings.CreateCategoryListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureCategoryListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneLabel('categoryManagementLabel', {
                  text: strings.CategoryManagementLabel
                }),
                PropertyPaneDropdown('selectedCategoryKey', {
                  label: strings.CategoryItemsFieldLabel,
                  options: categoryDropdownOptions,
                  selectedKey: this._categoryOptions.length > 0
                    ? this.properties.selectedCategoryKey
                    : '__no_categories__',
                  disabled: !canMutateCategories || this._categoryOptions.length === 0
                }),
                PropertyPaneTextField('newCategoryTitle', {
                  label: strings.NewCategoryFieldLabel,
                  value: this.properties.newCategoryTitle ?? '',
                  placeholder: strings.NewCategoryFieldPlaceholder,
                  disabled: !canMutateCategories
                }),
                PropertyPaneButton('addCategoryButton', {
                  text: '+',
                  ariaLabel: strings.AddCategoryButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddCategoryClick,
                  disabled: !canAddCategory
                }),
                PropertyPaneButton('removeCategoryButton', {
                  text: '-',
                  ariaLabel: strings.RemoveCategoryButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this._handleRemoveCategoryClick,
                  disabled: !canRemoveCategory
                }),
                PropertyPaneLabel('categoryStatus', {
                  text: this._categoryStatusMessage ?? ''
                }),
                PropertyPaneDropdown('statusListTitle', {
                  label: strings.StatusListFieldLabel,
                  options: [
                    { key: '', text: strings.StatusListDefaultOptionLabel },
                    ...this._listOptions
                  ],
                  selectedKey: this._selectedStatusListTitle ?? '',
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createStatusListButton', {
                  text: strings.CreateStatusListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureStatusListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneLabel('statusManagementLabel', {
                  text: strings.StatusManagementLabel
                }),
                PropertyPaneDropdown('selectedStatusKey', {
                  label: strings.StatusItemsFieldLabel,
                  options: statusDropdownOptions,
                  selectedKey: this._statusOptions.length > 0
                    ? this.properties.selectedStatusKey
                    : '__no_statuses__',
                  disabled: !canMutateStatuses || this._statusOptions.length === 0
                }),
                PropertyPaneTextField('newStatusTitle', {
                  label: strings.NewStatusFieldLabel,
                  value: this.properties.newStatusTitle ?? '',
                  placeholder: strings.NewStatusFieldPlaceholder,
                  disabled: !canMutateStatuses
                }),
                PropertyPaneButton('addStatusButton', {
                  text: '+',
                  ariaLabel: strings.AddStatusButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddStatusClick,
                  disabled: !canAddStatus
                }),
                PropertyPaneButton('removeStatusButton', {
                  text: '-',
                  ariaLabel: strings.RemoveStatusButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this._handleRemoveStatusClick,
                  disabled: !canRemoveStatus
                }),
                PropertyPaneLabel('statusStatus', {
                  text: this._statusStatusMessage ?? ''
                }),
                PropertyPaneDropdown('defaultStatus', {
                  label: strings.DefaultStatusFieldLabel,
                  options:
                    defaultStatusOptions.length > 0
                      ? defaultStatusOptions
                      : [{ key: defaultStatus, text: defaultStatus }],
                  selectedKey: defaultStatus,
                  disabled: defaultStatusOptions.length === 0
                }),
                PropertyPaneDropdown('subcategoryListTitle', {
                  label: strings.SubcategoryListFieldLabel,
                  options: [
                    { key: '', text: strings.SubcategoryListNoneOptionLabel },
                    ...this._listOptions
                  ],
                  selectedKey: this._selectedSubcategoryListTitle ?? '',
                  disabled: this._isLoadingLists && this._listOptions.length === 0
                }),
                PropertyPaneButton('createSubcategoryListButton', {
                  text: strings.CreateSubcategoryListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleEnsureSubcategoryListClick,
                  disabled: this._isListCreationInProgress
                }),
                PropertyPaneLabel('subcategoryManagementLabel', {
                  text: strings.SubcategoryManagementLabel
                }),
                PropertyPaneDropdown('selectedSubcategoryKey', {
                  label: strings.SubcategoryItemsFieldLabel,
                  options: subcategoryDropdownOptions,
                  selectedKey: this._subcategoryOptions.length > 0
                    ? this.properties.selectedSubcategoryKey
                    : '__no_subcategories__',
                  disabled: !canMutateSubcategories || this._subcategoryOptions.length === 0
                }),
                PropertyPaneTextField('newSubcategoryTitle', {
                  label: strings.NewSubcategoryFieldLabel,
                  value: this.properties.newSubcategoryTitle ?? '',
                  placeholder: strings.NewSubcategoryFieldPlaceholder,
                  disabled: !canMutateSubcategories
                }),
                PropertyPaneButton('addSubcategoryButton', {
                  text: '+',
                  ariaLabel: strings.AddSubcategoryButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleAddSubcategoryClick,
                  disabled: !canAddSubcategory
                }),
                PropertyPaneButton('removeSubcategoryButton', {
                  text: '-',
                  ariaLabel: strings.RemoveSubcategoryButtonAriaLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this._handleRemoveSubcategoryClick,
                  disabled: !canRemoveSubcategory
                }),
                PropertyPaneLabel('subcategoryStatus', {
                  text: this._subcategoryStatusMessage ?? ''
                }),
                PropertyPaneLabel('createListStatus', {
                  text: this._listCreationMessage ?? ''
                }),
                PropertyPaneDropdown('completedStatus', {
                  label: strings.CompletedStatusFieldLabel,
                  options:
                    completedStatusOptions.length > 0
                      ? completedStatusOptions
                      : [{ key: completedStatus, text: completedStatus }],
                  selectedKey: completedStatus,
                  disabled: completedStatusOptions.length === 0
                }),
                PropertyPaneTextField('headerTitle', {
                  label: strings.HeaderTitleFieldLabel,
                  value: this.properties.headerTitle
                }),
                PropertyPaneTextField('headerSubtitle', {
                  label: strings.HeaderSubtitleFieldLabel,
                  value: this.properties.headerSubtitle
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('useTableLayout', {
                  label: strings.UseTableLayoutToggleLabel,
                  onText: strings.UseTableLayoutToggleOnText,
                  offText: strings.UseTableLayoutToggleOffText
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
