declare interface IVivaOrgstructureWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  AdvancedGroupName: string;
  EmployeeListLinkFieldLabel: string;
  StartLoginFieldLabel: string;
  LoginDomainFieldLabel: string;
}

declare module 'VivaOrgstructureWebPartStrings' {
  const strings: IVivaOrgstructureWebPartStrings;
  export = strings;
}
