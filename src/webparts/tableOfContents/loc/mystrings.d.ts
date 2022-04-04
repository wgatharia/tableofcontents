declare interface ITableOfContentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  TocSourceFieldLabel: string;
  SortByFieldLabel: string;
  SortOrderLabel: string;
  RowLimitLabel: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
