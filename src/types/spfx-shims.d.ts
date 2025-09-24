declare module '@ms/odsp-core-bundle' {
  export interface IInternalAadTokenProvider {
    getToken(resourceEndpoint: string, useCachedToken?: boolean): Promise<string>;
  }
}

declare module '@microsoft/microsoft-graph-client' {
  export interface ClientOptions {
    authProvider?: unknown;
  }
  export interface Client {
    api(path: string): Client;
    version(version: string): Client;
    select(properties: string): Client;
    get(): Promise<unknown>;
  }
  export function Client(options?: ClientOptions): Client;
}

declare module '@ms/odsp-datasources/lib/interfaces/ISpPageContext' {
  export interface ISpPageContext {}
}

declare type AzureActiveDirectoryInfo = unknown;
declare type O365GroupAssociation = unknown;
declare type IPropertyPaneConsumer = unknown;
