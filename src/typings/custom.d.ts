declare module '@microsoft/sp-core-library';
declare module '@microsoft/sp-property-pane';
declare module '@microsoft/sp-webpart-base';
declare module '@microsoft/sp-lodash-subset';

declare module '*.module.scss' {
  const classes: { [key: string]: string };
  export default classes;
}

