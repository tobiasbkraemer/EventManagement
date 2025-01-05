import { WebPartContext } from "@microsoft/sp-webpart-base"; 
import { spfi } from '@pnp/sp';

export interface IMeetingManagementProps {
  context: WebPartContext;
  checkbox: boolean;
  registerOptions: string;
  filter: string;
  allListsArray: Array<any>;
  list: string;
  siteUrl: string;
  //showCustomEvents: boolean;
  selectedList: string;
  meetingsListOption: string;
  meetingRegistrationListOption: string;
  toggleRegistrations: boolean;
  toggleUserMenu: boolean;
  togglePagination: boolean;
  _width: number;
  seeAllUrl: string;
  numberOfMeetings: number;
  filterById: number;
  filterByMeetingType: string;
  filterByCategory: string;
  filterByRoom: string;
  sp: ReturnType<typeof spfi>; 
  types: Array<{ key: string, text: string }>;
  meetingCategories: Array<{ key: string, text: string }>;
  rooms: Array<{ key: string, text: string }>;
}
