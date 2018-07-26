import { IMeeting } from '.';

export interface IReactGraphdataState {
  error: string;
  loading: boolean;
  meetings: IMeeting[];
  renderedDateTime: Date;
  UsedFiles: any[];
}