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
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SamverkansportalenWebPartStrings';
import Samverkansportalen from './components/Samverkansportalen';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, ISamverkansportalenProps } from './components/ISamverkansportalenProps';
import GraphSuggestionsService from './services/GraphSuggestionsService';

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
  newListTitle?: string;
}

export default class SamverkansportalenWebPart extends BaseClientSideWebPart<ISamverkansportalenWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [
    { key: DEFAULT_SUGGESTIONS_LIST_TITLE, text: DEFAULT_SUGGESTIONS_LIST_TITLE }
  ];
  private _isLoadingLists: boolean = false;
  private _isCreatingList: boolean = false;
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
        listTitle: this._selectedListTitle
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

  private _handleCreateListClick = (): void => {
    void this._createListFromPropertyPane();
  };

  private async _createListFromPropertyPane(): Promise<void> {
    const rawTitle: string = (this.properties.newListTitle ?? '').trim();

    if (!rawTitle) {
      this._setListCreationMessage(strings.CreateListNameMissingMessage);
      return;
    }

    this._isCreatingList = true;
    this._setListCreationMessage(strings.CreateListProgressMessage);

    let message: string | undefined;

    try {
      const result: { created: boolean } = await this._getGraphService().ensureList(rawTitle);

      this.properties.listTitle = rawTitle;
      this.properties.newListTitle = '';
      this._addListOption(rawTitle);
      this.render();

      message = result.created
        ? strings.CreateListSuccessMessage.replace('{0}', rawTitle)
        : strings.CreateListAlreadyExistsMessage;
    } catch (error) {
      console.error('Failed to create the SharePoint list from the property pane.', error);
      message = strings.CreateListErrorMessage;
    } finally {
      this._isCreatingList = false;
      this._setListCreationMessage(message);
    }
  }

  private _setListCreationMessage(message?: string): void {
    this._listCreationMessage = message;
    this.context.propertyPane.refresh();
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private get _selectedListTitle(): string {
    return this._normalizeListTitle(this.properties.listTitle);
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
                PropertyPaneTextField('newListTitle', {
                  label: strings.NewListNameFieldLabel
                }),
                PropertyPaneButton('createListButton', {
                  text: strings.CreateListButtonLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._handleCreateListClick,
                  disabled: this._isCreatingList
                }),
                PropertyPaneLabel('createListStatus', {
                  text: this._listCreationMessage ?? ''
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
