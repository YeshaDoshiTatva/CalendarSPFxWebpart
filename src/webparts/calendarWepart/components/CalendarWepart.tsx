import * as React from 'react';
import styles from './CalendarWepart.module.scss';
import { ICalendarWepartProps } from './ICalendarWepartProps';
import { ICalendarItems } from '../Models/ICalendarItems';
import { ICalendarWepartState } from './ICalendarWebpartState';
import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICalendarWepartWebPartProps } from '../CalendarWepartWebPart';
// import {SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-core-library';
import * as strings from 'CalendarWepartWebPartStrings';
import {Icon} from 'office-ui-fabric-react';

export default class CalendarWepart extends React.Component<ICalendarWepartProps, ICalendarWepartState> {

  constructor(props: ICalendarWepartProps){
    super(props);

    this.state = {
      CalendarItems : [],
      ListURL: this.props.listUrl      
    }
  }

  private lstCalendarItem : ICalendarItems[];

  componentDidMount = () => {
    ///<summary>On load event.</summary>
    this.getListItems();
  };


  componentDidUpdate = (prevProps) => {
    ///<summary>Update event.</summary>
    /// <param name="prevProps">Previous Properties</param>
    if(prevProps.listUrl !== this.props.listUrl){
      if(this.props.listUrl !== undefined || this.props.listUrl !== null || this.props.listUrl.trim() !== ""){
        this.getListItems();
      }
    }
    if(prevProps.displayItems !== this.props.displayItems){
      this.getListItems();
    }
  };

    getListItems = () => {
    ///<summary>Get items from the list.</summary>
    if(this.props.listUrl !== undefined || this.props.displayItems !== undefined){
      let listURL = this.props.listUrl;
      let strCurrentURL = new URL(listURL);
      let currentCalendarAbsoluteURL = listURL.substr(0, listURL.lastIndexOf('/Lists/'));
      let pathname = strCurrentURL.pathname;
      let today = new Date().toISOString();
      
      this.props.spHttpClient.get(currentCalendarAbsoluteURL+"/_api/Web/GetList('"+pathname+"')/items?$select=ID,Title,Description,EventDate&$top="+this.props.displayItems+"&$filter=EventDate ge '"+today+"'", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if(response.ok){
          response.json().then((responseJSON) => {
            this.lstCalendarItem = responseJSON.value;
            if(this.lstCalendarItem != null && this.lstCalendarItem.length > 0){
               const messageDiv = document.querySelector("#divMessage");
               messageDiv.innerHTML = "";
               this.setState({CalendarItems: this.lstCalendarItem});
            }
            else if(this.lstCalendarItem.length === 0){
              const messageDiv = document.querySelector("#divMessage");
              messageDiv.innerHTML = strings.NoItemFoundMessage;
              this.setState({CalendarItems: []});
            }
          });
        }
        else{
          const messageDiv = document.querySelector("#divMessage");
          messageDiv.innerHTML = strings.NoItemFoundMessage;
          this.setState({CalendarItems: []});
        }
        
      }).catch((error) => {
        console.log("Error in getListItems ---->",error);
      });
    }
    else{
      const messageDiv = document.querySelector("#divMessage");
      messageDiv.innerHTML = strings.PropertiesMessage;
    }
  }
    
  public render(): React.ReactElement<ICalendarWepartProps> {
    ///<summary>Render method.</summary>
    return (
      <div className={ styles.calendarWepart }>
        <div className={ styles.container }>
          <div className={styles.clsMain}>
            <div className={styles.clsDivHeading}>
            <Icon iconName="Event" id={styles.icon} className='ms-Icon'/>
              <p className={styles.clsHeading}>Upcoming Events</p>
            </div>
            <p id="divMessage"></p>
            
            <div className="clsEvents">
                {this.state.CalendarItems.map(item => (
                  <div className={styles.msGridcol}>
                    <Icon iconName="EventInfo" id={styles.clsIcon} className="ms-Icon"/>
                     <h3><a href={this.props.listUrl+'/DispForm.aspx?ID='+item.ID}>{item.Title}</a></h3>
                     <p>{item.Description != null && item.Description.length > 0 ? item.Description.replace(/<[^>]+>/g, '') : ''}</p>
                  </div>
                ))}
            </div>
            </div>
        </div>
      </div>
    );
  }
}
