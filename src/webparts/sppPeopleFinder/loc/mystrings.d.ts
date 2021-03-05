declare interface ISppPeopleFinderWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SearchUser:string;
  SearchUserByFirstName:string;
  SearchUserByLastName:string;
  DisplayName:string;
  Email:string;
  MobilePhone:string;
  JobTitle:string;
  OfficeLocation:string;
  businessPhone:string;
  WebpartTitle:string;
  Title:string;
  placeholder:string;
}

declare module 'SppPeopleFinderWebPartStrings' {
  const strings: ISppPeopleFinderWebPartStrings;
  export = strings;
}