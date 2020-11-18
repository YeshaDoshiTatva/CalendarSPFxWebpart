import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarWepartWebPartStrings';
import CalendarWepart from './components/CalendarWepart';
import { ICalendarWepartProps } from './components/ICalendarWepartProps';
import { values } from 'office-ui-fabric-react';

export interface ICalendarWepartWebPartProps {
  listUrl: string;
  displayItems: string;
}

const arrNumberOfItems: IPropertyPaneDropdownOption[] = [{key:1,text:'1'},{key:2,text:'2'},{key:3,text:'3'},{key:4,text:'4'},{key:5,text:'5'}];

export default class CalendarWepartWebPart extends BaseClientSideWebPart <ICalendarWepartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarWepartProps> = React.createElement(
      CalendarWepart,
      {
        listUrl: this.properties.listUrl,
        displayItems: this.properties.displayItems,
        spHttpClient: this.context.spHttpClient,
        Title: "",
        Description: ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('listUrl', {
                  label: strings.ListURLFieldLabel,
                  value: ""
                }),
                PropertyPaneDropdown('displayItems', {
                  options: arrNumberOfItems,
                  label: strings.DisplayItems,
                  selectedKey: 5
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
