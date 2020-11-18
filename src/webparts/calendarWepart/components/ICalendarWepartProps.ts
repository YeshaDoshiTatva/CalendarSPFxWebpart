import { SPHttpClient } from '@microsoft/sp-http';

export interface ICalendarWepartProps {
  ListUrl: string;
  DisplayItems: string;
  spHttpClient: SPHttpClient;
  Title: string;
  Description : string;
}