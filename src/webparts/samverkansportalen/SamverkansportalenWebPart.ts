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
import { DEFAULT_SUGGESTIONS_LIST_TITLE, ISamverkansportalenProps } from './components/ISamverkansportalenProps';
import GraphSuggestionsService, {
  DEFAULT_SUBCATEGORY_LIST_TITLE
} from './services/GraphSuggestionsService';

type ListCreationType = 'suggestions' | 'votes' | 'subcategories';

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
  useTableLayout?: boolean;
  subcategoryListTitle?: string;
  voteListTitle?: string;
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

  public render(): void {
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
        useTableLayout: this.properties.useTableLayout,
        subcategoryListTitle: this._selectedSubcategoryListTitle
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
    this.properties.subcategoryListTitle = this._normalizeOptionalListTitle(this.properties.subcategoryListTitle);

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
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationStart(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ensureListOptions();
    this._listCreationMessage = undefined;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._extendConfiguredLists();
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
      ensureOption(this._selectedVoteListTitle);

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

  private _handleEnsureSubcategoryListClick = (): void => {
    this._ensureListFromPropertyPane('subcategories').catch(() => {
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
      } else {
        const result: { created: boolean } = await service.ensureSubcategoryList(trimmed);
        this.properties.subcategoryListTitle = trimmed;
        this._addListOption(trimmed);
        this.render();
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
      case 'subcategories':
        return strings.CreateListPromptSubcategoryLabel;
      default:
        return strings.CreateListPromptSuggestionsLabel;
    }
  }

  private _getDefaultListName(type: ListCreationType): string {
    switch (type) {
      case 'votes':
        return this._selectedVoteListTitle;
      case 'subcategories':
        return this._selectedSubcategoryListTitle ?? DEFAULT_SUBCATEGORY_LIST_TITLE;
      default:
        return this._selectedListTitle;
    }
  }

  private _getListProgressMessage(type: ListCreationType): string {
    switch (type) {
      case 'votes':
        return strings.CreateVotesListProgressMessage;
      case 'subcategories':
        return strings.CreateSubcategoryListProgressMessage;
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
    const subcategoryListTitle: string | undefined = this._selectedSubcategoryListTitle;

    try {
      await this._getGraphService().ensureList(listTitle);
      await this._getGraphService().ensureVoteList(voteListTitle);
      if (subcategoryListTitle) {
        await this._getGraphService().ensureSubcategoryList(subcategoryListTitle);
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

  private _normalizeOptionalListTitle(value?: string): string | undefined {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  private get _selectedSubcategoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.properties.subcategoryListTitle);
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
                PropertyPaneLabel('createListStatus', {
                  text: this._listCreationMessage ?? ''
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
