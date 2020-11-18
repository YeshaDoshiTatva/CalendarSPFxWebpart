import { SPHttpClient } from '@microsoft/sp-http';

export interface ICalendarWepartProps {
  listUrl: string;
  displayItems: string;
  spHttpClient: SPHttpClient;
  Title: string;
  Description : string;
}