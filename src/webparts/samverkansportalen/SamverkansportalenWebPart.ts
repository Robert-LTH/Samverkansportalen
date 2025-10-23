import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SamverkansportalenWebPartStrings';
import Samverkansportalen from './components/Samverkansportalen';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, ISamverkansportalenProps } from './components/ISamverkansportalenProps';

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
}

export default class SamverkansportalenWebPart extends BaseClientSideWebPart<ISamverkansportalenWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [
    { key: DEFAULT_SUGGESTIONS_LIST_TITLE, text: DEFAULT_SUGGESTIONS_LIST_TITLE }
  ];
  private _isLoadingLists: boolean = false;

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
  }

  private async _ensureListOptions(): Promise<void> {
    if (this._isLoadingLists) {
      return;
    }

    this._isLoadingLists = true;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title&$orderby=Title`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Unexpected response (${response.status}) while loading SharePoint lists.`);
      }

      const payload: { value: Array<{ Title?: string }> } = await response.json() as { value: Array<{ Title?: string }> };

      const options: IPropertyPaneDropdownOption[] = payload.value
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

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
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
