declare interface ISampleWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
   ListFieldLabel: string; // <-- Add this line
  
}

declare module 'SampleWebPartStrings' {
  const strings: ISampleWebPartStrings;
  export = strings;
}
