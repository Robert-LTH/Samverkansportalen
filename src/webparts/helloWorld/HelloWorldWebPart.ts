import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';

import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(HelloWorld, {
      sp: this._sp,
      currentUser: {
        loginName: this.context.pageContext.user.loginName || '',
        displayName: this.context.pageContext.user.displayName || 'Okänd användare'
      }
    });

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
      pages: []
    };
  }
}
