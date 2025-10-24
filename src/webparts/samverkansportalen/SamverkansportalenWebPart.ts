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
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SamverkansportalenWebPartStrings';
import Samverkansportalen from './components/Samverkansportalen';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, ISamverkansportalenProps } from './components/ISamverkansportalenProps';

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
  newListTitle?: string;
}

export default class SamverkansportalenWebPart extends BaseClientSideWebPart<ISamverkansportalenWebPartProps> {

  private static readonly LIST_REQUEST_ACCEPT_HEADERS: readonly string[] = [
    'application/json;odata=nometadata',
    'application/json;odata=minimalmetadata',
    'application/json;odata=verbose'
  ];

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [
    { key: DEFAULT_SUGGESTIONS_LIST_TITLE, text: DEFAULT_SUGGESTIONS_LIST_TITLE }
  ];
  private _isLoadingLists: boolean = false;
  private _isCreatingList: boolean = false;
  private _listCreationMessage?: string;

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
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listTitle: this._selectedListTitle
      }
    );

    ReactDom.render(element, this.domElement);
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
      const listUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title&$orderby=Title`;

      let lists: Array<{ Title?: string }> | undefined;

      for (const accept of SamverkansportalenWebPart.LIST_REQUEST_ACCEPT_HEADERS) {
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': accept
            }
          }
        );

        if (response.status === 406) {
          // Try the next Accept header variant for servers that do not understand this format.
          continue;
        }

        if (!response.ok) {
          throw new Error(`Unexpected response (${response.status}) while loading SharePoint lists.`);
        }

        const payload: unknown = await response.json();
        const parsedLists: Array<{ Title?: string }> | undefined = this._extractListItems(payload);

        if (parsedLists !== undefined) {
          lists = parsedLists;
          break;
        }
      }

      if (!lists) {
        throw new Error('Failed to load SharePoint lists because the response format was not recognized.');
      }

      const options: IPropertyPaneDropdownOption[] = lists
        .map((item) => item.Title)
        .filter((title): title is string => typeof title === 'string' && title.trim().length > 0)
        .map((title) => ({ key: title, text: title }))
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
      console.error('Failed to load available SharePoint lists for the property pane.', error);
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
      const result: 'created' | 'exists' = await this._ensureListExists(rawTitle);

      this.properties.listTitle = rawTitle;
      this.properties.newListTitle = '';
      this._addListOption(rawTitle);
      this.render();

      message = result === 'created'
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

  private async _ensureListExists(listTitle: string): Promise<'created' | 'exists'> {
    const listEndpoint: string = this._getListEndpoint(listTitle);

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      listEndpoint,
      SPHttpClient.configurations.v1,
      this._createOptions()
    );

    if (response.ok) {
      return 'exists';
    }

    if (response.status !== 404) {
      throw new Error(`Unexpected response (${response.status}) while checking for the ${listTitle} list.`);
    }

    await this._createListWithFields(listTitle);
    return 'created';
  }

  private async _createListWithFields(listTitle: string): Promise<void> {
    const siteUrl: string = this.context.pageContext.web.absoluteUrl;

    const createListResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      this._createOptions({
        Title: listTitle,
        Description: 'Stores user suggestions and votes from the Samverkansportalen web part.',
        BaseTemplate: 100,
        AllowContentTypes: true
      })
    );

    if (!createListResponse.ok) {
      throw new Error('Failed to create the suggestions list.');
    }

    const listEndpoint: string = this._getListEndpoint(listTitle);

    await this._createField(listEndpoint, {
      Title: 'Details',
      FieldTypeKind: 3
    });

    await this._createField(listEndpoint, {
      Title: 'Votes',
      FieldTypeKind: 9,
      DefaultValue: '0'
    });

    await this._createField(listEndpoint, {
      Title: 'Status',
      FieldTypeKind: 6,
      Choices: {
        results: ['Active', 'Done']
      },
      DefaultValue: 'Active'
    });

    await this._createField(listEndpoint, {
      Title: 'Voters',
      FieldTypeKind: 3
    });
  }

  private async _createField(listEndpoint: string, definition: Record<string, unknown>): Promise<void> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${listEndpoint}/fields`,
      SPHttpClient.configurations.v1,
      this._createOptions(definition)
    );

    if (!response.ok) {
      throw new Error(`Failed to create field ${(definition.Title as string) || 'unknown'}.`);
    }
  }

  private _createOptions(body?: unknown, extraHeaders?: Record<string, string>): ISPHttpClientOptions {
    const headers: Record<string, string> = {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0'
    };

    if (body !== undefined) {
      headers['Content-type'] = 'application/json;odata=nometadata';
    }

    if (extraHeaders) {
      for (const key in extraHeaders) {
        if (Object.prototype.hasOwnProperty.call(extraHeaders, key)) {
          const value: string | undefined = extraHeaders[key];
          if (typeof value === 'string') {
            headers[key] = value;
          }
        }
      }
    }

    const options: ISPHttpClientOptions = {
      headers
    };

    if (body !== undefined) {
      options.body = JSON.stringify(body);
    }

    return options;
  }

  private _getListEndpoint(listTitle: string): string {
    const escapedTitle: string = listTitle.replace(/'/g, "''");
    return `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${escapedTitle}')`;
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private _extractListItems(payload: unknown): Array<{ Title?: string }> | undefined {
    if (!payload || typeof payload !== 'object') {
      return undefined;
    }

    const withValue = payload as { value?: unknown };
    if (Array.isArray(withValue.value)) {
      return withValue.value as Array<{ Title?: string }>;
    }

    const withVerbose = payload as { d?: { results?: unknown } };
    if (withVerbose.d && Array.isArray(withVerbose.d.results)) {
      return withVerbose.d.results as Array<{ Title?: string }>;
    }

    return undefined;
  }

  private get _selectedListTitle(): string {
    return this._normalizeListTitle(this.properties.listTitle);
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
