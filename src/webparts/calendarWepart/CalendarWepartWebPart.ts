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

export interface ICalendarWepartWebPartProps {
  PropertiesListUrl: string;
  PropertiesDisplayItems: string;
}

const arrNumberOfItems: IPropertyPaneDropdownOption[] = [{key:1,text:'1'},{key:2,text:'2'},{key:3,text:'3'},{key:4,text:'4'},{key:5,text:'5'}];

export default class CalendarWepartWebPart extends BaseClientSideWebPart <ICalendarWepartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarWepartProps> = React.createElement(
      CalendarWepart,
      {
        ListUrl: this.properties.PropertiesListUrl,
        DisplayItems: this.properties.PropertiesDisplayItems,
        spHttpClient: this.context.spHttpClient,
        Title: "",
        Description: ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  onDispose = () => {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

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
                PropertyPaneTextField('PropertiesListUrl', {
                  label: strings.ListURLFieldLabel,
                  value: ""
                }),
                PropertyPaneDropdown('PropertiesDisplayItems', {
                  options: arrNumberOfItems,
                  label: strings.DisplayItems,
                  //selectedKey: 5,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
