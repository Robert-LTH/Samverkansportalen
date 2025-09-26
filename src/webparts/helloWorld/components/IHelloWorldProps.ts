import { SPFI } from '@pnp/sp';

export interface ICurrentUserInfo {
  loginName: string;
  displayName: string;
}

export interface IHelloWorldProps {
  sp: SPFI;
  currentUser: ICurrentUserInfo;
}
