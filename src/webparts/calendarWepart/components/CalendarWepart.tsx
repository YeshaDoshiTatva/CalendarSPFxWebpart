import * as React from 'react';
import styles from './CalendarWepart.module.scss';
import { ICalendarWepartProps } from './ICalendarWepartProps';
import { ICalendarItems } from '../Models/ICalendarItems';
import { ICalendarWepartState } from './ICalendarWebpartState';
import { SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import * as strings from 'CalendarWepartWebPartStrings';
import {Icon} from 'office-ui-fabric-react';

export default class CalendarWepart extends React.Component<ICalendarWepartProps, ICalendarWepartState> {

  constructor(props: ICalendarWepartProps){
    super(props);

    this.state = {
      CalendarItems : [],
      ListURL: this.props.ListUrl   
    }
  }

  private lstCalendarItem : ICalendarItems[];

    componentDidMount = () => {
      ///<summary>On load event.</summary>
      this.getListItems();
    };

    componentDidUpdate = (prevProps) => {
      ///<summary>Event called when any states is changed.</summary>
      ///<param name="prevProps">Previous Properties values</param>
      if(prevProps.ListUrl !== this.props.ListUrl){
        if(this.props.ListUrl !== undefined || this.props.ListUrl !== null || this.props.ListUrl.trim() !== ""){
          this.getListItems();
        }
      }

      if(prevProps.DisplayItems !== this.props.DisplayItems){
        this.getListItems();
      }
    };

    getListItems = () => {
    ///<summary>Get items from the list.</summary>
    const messageDiv = document.querySelector("#divMessage");
    if(this.props.ListUrl !== undefined && this.props.DisplayItems !== undefined){
      let strListURLString = this.props.ListUrl;
      let strCurrentURL = new URL(strListURLString);
      let strListAbsoluteURL = strListURLString.substr(0, strListURLString.lastIndexOf('/Lists/'));
      let strListPathName = strCurrentURL.pathname;
      let todayDate = new Date().toISOString();
      
      this.props.spHttpClient.get(strListAbsoluteURL+"/_api/Web/GetList('"+strListPathName+"')/items?$select=ID,Title,Description,EventDate&$top="+this.props.DisplayItems+"&$filter=EventDate ge '"+todayDate+"'", SPHttpClient.configurations.v1)
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
          messageDiv.innerHTML = strings.NoItemFoundMessage;
          this.setState({CalendarItems: []});
        }
      }).catch((error) => {
        console.log("Error in getListItems ---->",error);
      });
    }
    else{
      messageDiv.innerHTML = strings.PropertiesMessage;
      this.setState({CalendarItems: []});
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
            <p id="divMessage" className={styles.clsMessage}></p>
            
            <div className="clsEvents">
                {this.state.CalendarItems.map(item => (
                  <div className={styles.msGridcol}>
                    <Icon iconName="EventInfo" id={styles.clsIcon} className="ms-Icon"/>
                     <h3><a target="_blank" href={this.props.ListUrl+'/DispForm.aspx?ID='+item.ID} className={styles.clsLink}>{item.Title}</a></h3>
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
