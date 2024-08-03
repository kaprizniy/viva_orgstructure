import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IVivaOrgstructureProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  employeeListLink: string;
  startLogin: string;
  loginDomain: string;
}

export interface IVivaOrgstructureState {
  orgUser: any;
  orgEmployees: any;
  loginName: string;
}