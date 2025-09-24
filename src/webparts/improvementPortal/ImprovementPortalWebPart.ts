import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ImprovementPortalWebPartStrings';
import ImprovementPortal, { IImprovementPortalProps } from './components/ImprovementPortal';
import { spfi, SPFI, SPFx } from '@pnp/sp';

export interface IImprovementPortalWebPartProps {
  description: string;
}

export default class ImprovementPortalWebPart extends BaseClientSideWebPart<IImprovementPortalWebPartProps> {

  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<IImprovementPortalProps> = React.createElement(
      ImprovementPortal,
      {
        description: this.properties.description,
        sp: this._sp
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
