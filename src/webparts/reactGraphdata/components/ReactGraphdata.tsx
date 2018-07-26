import * as React from 'react';
import styles from './ReactGraphdata.module.scss';
import * as strings from 'ReactGraphdataWebPartStrings';
//import { IReactGraphdataProps } from './IReactGraphdataProps';
import { IReactGraphdataProps, IMeeting, IMeetings, IReactGraphdataState } from '.';
//import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Link } from 'office-ui-fabric-react/lib/components/Link';

import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DocumentCardType } from 'office-ui-fabric-react/lib/components/DocumentCard';

import { escape } from '@microsoft/sp-lodash-subset';



export interface UsedDocumentsProp {
  id: string;
  lastUsed: LastUsed;
  fileurl: string;
  resourceReference: ResourceReference;
  resourceVisualization: ResourceVisualization;
}

export interface onedrivecreatedby{
  displayname:string;
  email:string;
  id:string;
}
export interface fileSystemInfo{
  createdDateTime:string;
  lastModifiedDateTime:string;
}

export interface OnedriveDataProp{
  id:string;
  createdby: onedrivecreatedby;
  createdDateTime: string;
  fileSystemInfo: fileSystemInfo;
  size:number;
  webUrl:string;
  downloadurl: string;
  name:string;
}

export interface ResourceReference {
  webUrl: string;
  id: string;
  type: string;
}

export interface ResourceVisualization {
  title: string;
  type: string;
  mediaType: string;
  previewImageUrl: string;
  previewText: string;
  containerWebUrl: string;
  containerDisplayName: string;
  containerType: string;
}

export interface LastUsed {
  lastAccessedDateTime: string;
  lastModifiedDateTime: string;
}

export default class ReactGraphdata extends React.Component<IReactGraphdataProps, IReactGraphdataState> {

  private _interval: number;
  constructor(props: IReactGraphdataProps) {
    super(props);
    this.state = {
      meetings: [],
      loading: false,
      error: undefined,
      renderedDateTime: new Date(),
      UsedFiles: []
    };
  }

  /**
   * Render meeting item
   */
  private _onRenderCell = (item: IMeeting, index: number | undefined): JSX.Element => {
    const startTime: Date = new Date(item.start.dateTime);
    const minutes: number = startTime.getMinutes();

    return <div className={`${styles.meetingWrapper} ${item.showAs}`}>
      <Link href={item.webLink} className={styles.meeting} target='_blank'>
        <div className={styles.start}>{`${startTime.getHours()}:${minutes < 10 ? '0' + minutes : minutes}`}</div>
        <div className={styles.subject}>{item.subject}</div>
        <div className={styles.duration}>{this._getDuration(item)}</div>
        <div className={styles.location}>{item.location.displayName}</div>
      </Link>
    </div>;
  }

  /**
   * Get user-friendly string that represents the duration of the meeting
   * < 1h: x minutes
   * >= 1h: 1 hour (y minutes)
   * all day: All day
   */
  private _getDuration = (meeting: IMeeting): string => {
    if (meeting.isAllDay) {
      return "All Days";
    }
    const startDateTime: Date = new Date(meeting.start.dateTime);
    const endDateTime: Date = new Date(meeting.end.dateTime);
    // get duration in minutes
    const duration: number = Math.round((endDateTime as any) - (startDateTime as any)) / (1000 * 60);
    if (duration <= 0) {
      return '';
    }

    if (duration < 60) {
      return `${duration} ${strings.Minutes}`;
    }

    const hours: number = Math.floor(duration / 60);
    const minutes: number = Math.round(duration % 60);
    let durationString: string = `${hours} ${hours > 1 ? strings.Hours : strings.Hour}`;
    if (minutes > 0) {
      durationString += ` ${minutes} ${strings.Minutes}`;
    }

    return durationString;
  }


  /**
   * Load recent messages for the current user
   */
  private _loadMeetings(): void {
    if (!this.props.graphClient) {
      return;
    }

    // update state to indicate loading and remove any previously loaded
    // meetings
    this.setState({
      error: null,
      loading: true,
      meetings: []
    });

    const date: Date = new Date();
    const now: string = date.toISOString();
    // set the date to midnight today to load all upcoming meetings for today
    date.setUTCHours(23);
    date.setUTCMinutes(59);
    date.setUTCSeconds(0);
    date.setDate(date.getDate() + (this.props.daysInAdvance || 0));
    const midnight: string = date.toISOString();

    this.props.graphClient
      // get all upcoming meetings for the rest of the day today
      .api(`me/calendar/calendarView?startDateTime=${now}&endDateTime=${midnight}`)
      //.api(`me/calendar/calendarView`)
      .version("v1.0")
      .select('subject,start,end,showAs,webLink,location,isAllDay')
      .top(this.props.numMeetings > 0 ? this.props.numMeetings : 100)
      // sort ascending by start time
      .orderby("start/dateTime")
      .get((err: any, res: IMeetings): void => {
        if (err) {
          // Something failed calling the MS Graph
          this.setState({
            error: err.message ? err.message : strings.Error,
            loading: false
          });
          return;
        }

        // Check if a response was retrieved
        if (res && res.value && res.value.length > 0) {
          this.setState({
            meetings: res.value,
            loading: false
          });
        }
        else {
          // No meetings found
          this.setState({
            loading: false
          });
        }
      });

    //this.props.graphClient.api(`me/insights/used`).version("beta")
    this.props.graphClient.api(`me/drive/root/children`).version("v1.0")
      .top(this.props.numMeetings > 0 ? this.props.numMeetings : 100)
      // sort ascending by start time
      //.orderby("start/dateTime")
      .get((err: any, res: any): void => {
        console.log(res);
        var mydata = res.value;


        let Itemdata = mydata.map(a => {
          let link: OnedriveDataProp = {
            id: a.id,
            fileSystemInfo: a.fileSystemInfo,
            size: a.size,
            webUrl: a.webUrl,
            createdby: a.createdBy.user,
            createdDateTime: a.createdDateTime,
            downloadurl:'',
            name: a.name
            //downloadurl: a.@microsoft.graph.downloadUrl
          };
          return link;
        });
        console.log(Itemdata);
        this.setState({ UsedFiles: Itemdata });

      });
  }

  /**
   * Sets interval so that the data in the component is refreshed on the
   * specified cycle
   */
  private _setInterval = (): void => {
    let { refreshInterval } = this.props;
    // set up safe default if the specified interval is not a number
    // or beyond the valid range
    if (isNaN(refreshInterval) || refreshInterval < 0 || refreshInterval > 60) {
      refreshInterval = 5;
    }
    // refresh the component every x minutes
    this._interval = setInterval(this._reRender, refreshInterval * 1000 * 60);
    this._reRender();
  }

  /**
   * Forces re-render of the component
   */
  private _reRender = (): void => {
    // update the render date to force reloading data and re-rendering
    // the component
    this.setState({ renderedDateTime: new Date() });
  }

  public componentDidMount(): void {
    this._setInterval();
  }

  public componentWillUnmount(): void {
    // remove the interval so that the data won't be reloaded
    clearInterval(this._interval);
  }

  public componentDidUpdate(prevProps: IReactGraphdataProps, prevState: IReactGraphdataState): void {
    // if the refresh interval has changed, clear the previous interval
    // and setup new one, which will also automatically re-render the component
    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      clearInterval(this._interval);
      this._setInterval();
      return;
    }

    // reload data on new render interval
    if (prevState.renderedDateTime !== this.state.renderedDateTime) {
      this._loadMeetings();
    }
  }

  public render(): React.ReactElement<IReactGraphdataProps> {
    console.log(this.state.UsedFiles);
    return (
      <div className={styles.personalCalendar}>
        {
          this.state.UsedFiles &&
            this.state.UsedFiles.length > 0 ? (
              <div>
                {this.state.UsedFiles.map(function (item, key) {
                  //console.log(item);
                  //return (<div itemID={item.Id}>{item.resourceVisualization.title}</div>)
                  /* This is for TrandingUsed API. const previewProps: IDocumentCardPreviewProps = {
                    previewImages: [
                      {
                        name: item.resourceVisualization.title,
                        url: item.resourceReference.webUrl,
                        previewImageSrc: item.resourceVisualization.previewImageUrl,
                        iconSrc: '',
                        imageFit: ImageFit.cover,
                        width: 318,
                        height: 196
                      }
                    ]
                  }; */

                  const previewProps: IDocumentCardPreviewProps = {
                    previewImages: [
                      {
                        name: item.name,
                        url: item.webUrl,
                        previewImageSrc: '',
                        iconSrc: '',
                        imageFit: ImageFit.cover,
                        width: 318,
                        height: 196
                      }
                    ]
                  };

                  return (
                    <DocumentCard onClickHref={item.webUrl} type={DocumentCardType.compact}>
                      {/* <DocumentCardPreview {...previewProps} /> */}
                      <DocumentCardTitle
                        title={item.name}
                        shouldTruncate={true}
                      />
                      <DocumentCardActivity
                        activity="Created a few minutes ago"
                        people={[{ name: item.createdby.displayName, profileImageSrc: `/_layouts/15/userphoto.aspx?size=M&username=${item.createdby.email}` }]}
                      />
                    </DocumentCard>
                  );
                })}

              </div>
            ) : (
              !this.state.loading && (
                this.state.error ?
                  <span className={styles.error}>{this.state.error}</span> :
                  <span className={styles.noMeetings}>No Meeting found.</span>
              )
            )
        }
      </div>
    );
  }
}
