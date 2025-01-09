import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneGroup,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import MeetingManagement from './components/MeetingManagement';
import { IMeetingManagementProps } from './components/IMeetingManagementProps';
import { spfi, SPFx } from "@pnp/sp";
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/content-types/list";
import "@pnp/sp/lists/web";
import "@pnp/sp/sites";
import * as strings from 'MeetingManagementWebPartStrings';



export interface IMeetingManagementWebPartProps {
  checkbox: boolean;
  selectedList: string;
  registerOptions: string;
  filter: string;
  allListsArray: Array<any>;
  list: string;
  siteUrl: string;
  meetingsListOption: string;
  meetingRegistrationListOption: string;
  toggleRegistrations: boolean;
  toggleUserMenu: boolean;
  togglePagination: boolean;
  _width: number;
  seeAllUrl: string;
  numberOfMeetings: number;
  backgroundColorOptions: string;
  settings: any;
  filterById: number;
  filterByMeetingType: string;
  filterByCategory: string;
  filterByRoom: string;
  meetingTypes: Array<{ key: string, text: string }>;
  categories: Array<{ key: string, text: string }>;
  rooms: Array<{ key: string, text: string }>;
  sp: ReturnType<typeof spfi>;

  //to refactor later (temporary solution)
  /* types: Array<{ key: string, text: string }>;
  meetingCategories: Array<{ key: string, text: string }>;
  meetingRooms: Array<{ key: string, text: string }>; */
}

export default class MeetingManagementWebPart extends BaseClientSideWebPart<IMeetingManagementWebPartProps> {



  public render(): void {

    const element: React.ReactElement<IMeetingManagementProps> = React.createElement (


      MeetingManagement,
      {
        context: this.context,
        checkbox: this.properties.checkbox,
        selectedList: this.properties.selectedList,
        registerOptions: this.properties.registerOptions,
        filter: this.properties.filter,
        allListsArray: this.properties.allListsArray,
        siteUrl: this.properties.siteUrl,
        list: this.properties.list,
        meetingsListOption: this.properties.meetingsListOption,
        meetingRegistrationListOption: this.properties.meetingRegistrationListOption,
        toggleRegistrations: this.properties.toggleRegistrations,
        toggleUserMenu: this.properties.toggleUserMenu,
        togglePagination: this.properties.togglePagination,
        _width: this.properties._width,
        seeAllUrl: this.properties.seeAllUrl,
        numberOfMeetings: this.properties.numberOfMeetings,
        filterById: this.properties.filterById,
        filterByMeetingType: this.properties.filterByMeetingType,
        filterByCategory: this.properties.filterByCategory,
        filterByRoom: this.properties.filterByRoom,
        sp: this.properties.sp,
        types: this.properties.meetingTypes,
        meetingCategories: this.properties.categories,
        rooms: this.properties.rooms
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    console.log("onInit called");
    await super.onInit();


    this.properties.sp = spfi().using(SPFx(this.context));

    this.properties.filterByMeetingType = this.properties.filterByMeetingType || "";
    this.properties.filterByCategory = this.properties.filterByCategory || "";
    this.properties.filterByRoom = this.properties.filterByRoom || "";


    this.properties._width = this.width;


    await this.getEventLists(this.properties.siteUrl);

    if (this.properties.meetingsListOption) {
      await this.getMeetingChoices(this.properties.meetingsListOption);
    }

    this.render();
  }



  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }


  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    console.log(`Property changed: ${propertyPath}, Old Value: ${oldValue}, New Value: ${newValue}`);

    if (propertyPath === 'siteUrl' && newValue !== oldValue) {
      console.log("Site URL changed");
      this.properties.siteUrl = newValue;
      await this.getEventLists(newValue);
      this.context.propertyPane.refresh();
    }

    if (['meetingsListOption', 'meetingRegistrationListOption'].includes(propertyPath) && newValue !== oldValue) {
      console.log("List option changed for: ", propertyPath);
      this.properties[propertyPath] = newValue;

      if (propertyPath === 'meetingsListOption') {
        await this.getMeetingChoices(newValue);
      }

      this.context.propertyPane.refresh();
      this.render();
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log("Property Pane Configuration called");

    let settingsGroups: IPropertyPaneGroup[] = [];

    try {
      this.properties.meetingTypes = this.properties.meetingTypes || [];
      this.properties.categories = this.properties.categories || [];
      this.properties.rooms = this.properties.rooms || [];




      const contentSettingsGroupName = [
        PropertyPaneTextField("siteUrl", {
          label: strings.SiteURLLabel,
          description: strings.SiteURLDescription,
          placeholder: "https://example.sharepoint.com/sites/exampleSite",
          onGetErrorMessage: (value) => this.getEventLists(value),
        }),
        PropertyPaneDropdown("meetingsListOption", {
          label: strings.MeetingsListOption,
          options: this.properties.allListsArray,
          selectedKey: this.properties.meetingsListOption,
        }),
        PropertyPaneDropdown("meetingRegistrationListOption", {
          label: strings.MeetingRegistrationListOption,
          options: this.properties.allListsArray,
          selectedKey: this.properties.meetingRegistrationListOption,
        }),
        PropertyPaneTextField("seeAllUrl", {
          label: strings.SeeAllUrlLabel,
          placeholder: "https://pnmmhdev.sharepoint.com/sites/exampleSite/Lists/Meetings/AllItems.aspx",
        }),
        PropertyPaneDropdown("filterByMeetingType", {
          label: strings.FilterByMeetingType,
          options: [
            { key: "", text: strings.All },
            ...(
              this.properties.meetingTypes.length > 0
                ? this.properties.meetingTypes.map(type => ({
                  key: type.key,
                  text: type.text
                }))
                : [{ key: '', text: 'No Meeting Types Available' }]
            )
          ],
          selectedKey: this.properties.filterByMeetingType,
        }),



        PropertyPaneDropdown("filterByCategory", {
          label: strings.FilterByCategory,
          options: [
            { key: "", text: strings.All },
            ...(
              this.properties.categories.length > 0
                ? this.properties.categories.map(category => ({
                  key: category.key,
                  text: category.text
                }))
                : [{ key: '', text: 'No Categories Available' }]
            )
          ],
          selectedKey: this.properties.filterByCategory,
        }),

        PropertyPaneDropdown("filterByRoom", {
          label: strings.FilterByRoom,
          options: [
            { key: "", text: strings.All },
            ...(
              this.properties.rooms.length > 0
                ? this.properties.rooms.map(room => ({
                  key: room.key,
                  text: room.text
                }))
                : [{ key: '', text: 'No Rooms Available' }]
            )
          ],
          selectedKey: this.properties.filterByRoom,
        }),

      ];

      const detailsSettingsGroupName = [
        PropertyPaneDropdown("registerOptions", {
          label: strings.RegisterOptionsLabel,
          options: [
            { key: "showInView", text: strings.ShowInView },
            { key: "showInEvent", text: strings.ShowInEvent },
            { key: "showBoth", text: strings.ShowBoth },
            { key: "hide", text: strings.Hide },
          ],
        }),
        PropertyPaneToggle("toggleRegistrations", {
          label: strings.ShowRegistrationsLabel,
          onText: strings.ToggleOn,
          offText: strings.ToggleOff,
          onAriaLabel: strings.ToggleOn,
          offAriaLabel: strings.ToggleOff,

        }),
        PropertyPaneToggle("toggleUserMenu", {
          label: strings.ShowUserMenuLabel,
          onText: strings.ToggleOn,
          offText: strings.ToggleOff,
          onAriaLabel: strings.ToggleOn,
          offAriaLabel: strings.ToggleOff,

        }),
        PropertyPaneSlider('numberOfMeetings', {
          label: strings.NumberOfMeetings,
          min: 1,
          max: 10,
          value: this.properties.numberOfMeetings,
          showValue: true
        }),

      ];

      settingsGroups.push(
        {
          groupName: strings.ContentSettingsGroupName,
          groupFields: contentSettingsGroupName,
          isCollapsed: false,
        },
        {
          groupName: strings.DetailsSettingsGroupName,
          groupFields: detailsSettingsGroupName,
          isCollapsed: false,
        },

      );

    } catch (error) {
      console.log(error);
    }

    return {
      pages: [
        {
          header: {
            description: "Her kan du tilpasse indhold og visning.",
          },
          displayGroupsAsAccordion: true,
          groups: settingsGroups,
        },
      ],
    };
  }

 



  // Get all Event lists from the site
  private getEventLists = async (value?: string): Promise<string> => {
    try {
      const listContentType1 = "0x010200"; // Events list
      const listContentType2 = "0x012000"; // Custom list
      const listContentType3 = "0x01001"; // Custom list

      const sp = spfi().using(SPFx(this.context));
      const web = value ? Web([sp.web, value]) : sp.site.rootWeb;


      web.lists().then((allLists) => {
        let correctContentTypeLists: any[] = [];
        let promiseArray = allLists.map((list) => {
          return new Promise((resolve, reject) => {
            web.lists.getById(list.Id)
              .contentTypes()
              .then((item) => {
                let contentTypeId = item[item.length - 1].Id.StringValue;
                if (contentTypeId && (contentTypeId.indexOf(listContentType2) > -1 || contentTypeId.indexOf(listContentType1) > -1) || contentTypeId.indexOf(listContentType3) > -1) {
                  correctContentTypeLists.push({
                    key: list.Id,
                    text: list.Title,
                    contentTypeId: item[item.length - 1].Id.StringValue,
                  });
                }
                resolve(true);
              }).catch((error) => {
                reject(error);
                this.properties.allListsArray = [];
                return "URL not correct";
              });
          });
        });

        Promise.all(promiseArray).then((res) => {
          this.properties.allListsArray = correctContentTypeLists.concat();
          if (this.context != undefined) {
            this.context.propertyPane.refresh();
          }
        });
      });
      return null;
    } catch (error) {
      this.properties.allListsArray = [];
      console.log(error);
    }
  }

  private async getMeetingChoices(listId: string): Promise<void> {
    try {
      // Hent listen med de nødvendige felter
      const listFileds = await this.properties.sp.web.lists.getById(listId).fields
        .filter(`InternalName eq 'MeetingType' or InternalName eq 'Room' or InternalName eq 'Category'`)
        .select("Choices", "InternalName", "odata.type")(); // Retrieve all list fields

      const meetingTypes = listFileds.find(field => field.InternalName === 'MeetingType').Choices;
      const rooms = listFileds.find(field => field.InternalName === 'Room').Choices;
      const categories = listFileds.find(field => field.InternalName === 'Category').Choices;

      // Tildel de distinkte værdier til egenskaberne i this.properties
      this.properties.meetingTypes = meetingTypes.map(type => ({ key: type, text: type }));
      this.properties.rooms = rooms.map(room => ({ key: room, text: room }));
      this.properties.categories = categories.map(category => ({ key: category, text: category }));

      // Hvis der skal opdateres UI, kan propertyPane opdateres her
      if (this.context.propertyPane) {
        this.context.propertyPane.refresh();
      }

      console.log("Room Choices Updated: ", this.properties.meetingTypes, this.properties.rooms, this.properties.categories);


    } catch (error) {
      console.error('Error fetching Meeting Choices:', error);
    }
  }



}
