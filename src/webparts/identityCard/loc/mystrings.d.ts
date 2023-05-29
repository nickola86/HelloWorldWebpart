declare interface IIdentityCardWebPartStrings {
  DettaglioUtentePaneDescription: string;
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

  CognomeFieldLabel: string;
  NomeFieldLabel: string;
  LuogoDiNascitaFieldLabel: string;
  GenereFieldLabel: string;
  DataDiNascitaFieldLabel: string;
  ImmagineBase64FieldLabel: string;

  DateFormatErrorMessage: string;
}

declare module 'IdentityCardWebPartStrings' {
  const strings: IIdentityCardWebPartStrings;
  export = strings;
}
