import * as React from 'react';
import { IMeetingManagementProps } from './IMeetingManagementProps';
import { IMeeting } from './IMeeting';
import { IMeetingManagementState } from './IMeetingManagementState';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/search';
import { CommandButton, Shimmer, ShimmerElementType } from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';
import { Icon } from 'office-ui-fabric-react';
import MeetingManagementCell from './MeetingManagementCell';
import MeetingManagementDropdowns from './MeetingManagementDropdowns';
import MeetingManagementDetails from './MeetingManagementDetails';
import MeetingManagementConfirmation from './MeetingManagementConfirmation';
import MeetingManagementSuccess from './MeetingManagementSuccess';
import styles from './MeetingManagement.module.scss';
import * as strings from 'MeetingManagementWebPartStrings'


export default class MeetingManagement extends React.Component<IMeetingManagementProps, IMeetingManagementState> {


  private registrationList: any;

  private currentTime: Date = new Date();

  constructor(props: IMeetingManagementProps) {
    super(props);

    this.state = {
      meetings: [],
      showMeetingData: false,
      showMeetings: true,
      selectedMeetingId: 0,
      isDialogVisible: false,
      selectedMeeting: null,
      selectedTypeFilter: "",
      selectedCategoryFilter: "",
      selectedRoomFilter: "",
      registeredMeetingIds: [],
      registrationsCount: 0,
      isUserRegistered: false,
      registeredUserNames: [],
      isCalloutVisible: false,
      showMyRegistrations: false,
      dataLoaded: false,
      isEditor: false,
      showRegistrations: false,
      currentMeetingPage: 1,
      meetingsPerPage: 0,
      layout: "",
      meetingsCount: 0,
      showPagination: false,
      registeredUsers: [],
      registrationCounts: [],
      userId: 0,
      userName: "",
      userEmail: "",
      confirmationDialogVisible: false,
      confirmationDialogMessage: '',
      successDialogVisible: false,
      successDialogMessage: '',
      confirmationDialogAction: null,
      requiresApproval: false,
      approvalMessage: "",
      confirmationDialogTitle: "",
      confirmationDialogButtonText: "",
      successDialogTitle: "",
      successDialogButtonText: "",
      allRegistrations: []
    };

    this.showConfirmationDialog = this.showConfirmationDialog.bind(this);
    this.registerForMeeting = this.registerForMeeting.bind(this);
    this.unregisterFromMeeting = this.unregisterFromMeeting.bind(this);

  }

  public render(): React.ReactElement<IMeetingManagementProps> {

    // Shimmer
    if (!this.state.dataLoaded) {
      return this.renderShimmer();
    }

    // Get the filtered meetings
    let filteredMeetings = this.filterMeetings();

    // Ensure the number of meetings to display is based on props
    const totalMeetingsToShow = Math.min(filteredMeetings.length, this.props.numberOfMeetings || 4);
    filteredMeetings = filteredMeetings.slice(0, totalMeetingsToShow);

    // Pagination logic: adjust the slicing based on current page and meetings per page
    /* const startIndex = (this.state.currentMeetingPage - 1) * this.state.meetingsPerPage;
    const endIndex = Math.min(startIndex + this.state.meetingsPerPage, filteredMeetings.length);
 */
    // Get the paginated meetings
    const paginatedMeetings = this.getPaginatedMeetings(filteredMeetings);

    // Determine how many items are on the current page
    const itemsOnCurrentPage = paginatedMeetings.length;

    // Get the justify class based on layout and items on current page 
    const justifyClass = this.getJustifyClass(itemsOnCurrentPage);


    return (
      <div>
        <h1
          className={styles.overskrift}
          aria-label={strings.MeetingManagementHeader}
        >
          {strings.MeetingManagementHeader}
        </h1>

        {this.props.toggleUserMenu && (
          <div>
            {/* Dropdowns and toggle */}
            <MeetingManagementDropdowns
              meetingTypes={this.props.types}
              categories={this.props.meetingCategories}
              rooms={this.props.rooms}
              onTypeFilterChange={(selectedType) => this.setState({ selectedTypeFilter: selectedType })}
              onCategoryFilterChange={(selectedCategory) => this.setState({ selectedCategoryFilter: selectedCategory })}
              onRoomFilterChange={(selectedRoom) => this.setState({ selectedRoomFilter: selectedRoom })}
              onToggleChange={(checked) => this.handleToggleMyRegistrations(checked)}
              selectedTypeFilter={this.state.selectedTypeFilter}
              selectedCategoryFilter={this.state.selectedCategoryFilter}
              selectedRoomFilter={this.state.selectedRoomFilter}
              showMyRegistrations={this.state.showMyRegistrations}
              filterByMeetingType={this.props.filterByMeetingType}
              filterByCategory={this.props.filterByCategory}
              filterByRoom={this.props.filterByRoom}
            />
          </div>
        )}

        {/* Add and see-all for meetings */}
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          {this.state.isEditor && (
            <div className={styles.addMeetingButton}>
              <CommandButton
                iconProps={{
                  iconName: 'Add',
                  className: styles.textInherit
                }}
                text={"Tilføj møde"}
                styles={{
                  root: { color: 'inherit' },
                }}
                onClick={ev => window.open("https://pnmmhdev.sharepoint.com/sites/Tobias/_layouts/15/listform.aspx?PageType=8&ListId=%7B889DCCCC-8E8A-43E4-A076-6D72D1FB66DD%7D&RootFolder=&Source=https%3A%2F%2Fpnmmhdev.sharepoint.com%2Fsites%2FTobias%2FLists%2FMeetings%2FAllItems.aspx%3Fviewid%3D85302bd0%252D8bb9%252D4f68%252D9f48%252D005e11fdacbd%26npsAction%3DcreateList&ContentTypeId=0x0100F7E6163A74021E4299660258EAFF215600159C9AB3B42A3A49A1BDA6B677DE1D7F", "_blank")}
              />
            </div>
          )}


          <CommandButton
            text={strings.SeeAll}
            onClick={this.handleSeeAllClick}
            className={styles.seeAll}
            styles={{
              root: { color: 'inherit' },
            }}
          />
        </div>


        {/* Render the meetings and arrows */}
        <div className={`${styles.cellContainer} ${styles[this.state.layout]} ${justifyClass}`}>


          {/* Left Arrow - visible when more than one page */}
          {this.hasMultiplePages() && (
            <button
              className={`${styles.arrows} ${styles.arrowLeft}`}
              onClick={this.handlePreviousMeetingPage}
              data-action="tabAction"
              onKeyDown={(event) => { if (event.key == "Enter") this.handlePreviousMeetingPage }}
            />
          )}


          {/* Render the meetings */}
          {paginatedMeetings.map((meeting) => (
            <MeetingManagementCell
              key={meeting.Id}
              meeting={meeting}
              layout={this.state.layout}
              registrationCounts={this.state.registrationCounts}
              registeredMeetingIds={this.state.registeredMeetingIds}
              currentDate={this.currentTime}
              showMeetingDetails={this.showMeetingDetails}
              registerForMeeting={this.registerForMeeting}
              unregisterFromMeeting={this.unregisterFromMeeting}
              justifyClass={this.getJustifyClass(itemsOnCurrentPage)}
              registerOptions={this.props.registerOptions}
              responsibleUser={{
                userName: this.state.userName,
                userEmail: this.state.userEmail
              }}
              getUserProfilePicture={this.getUserProfilePicture}
            />
          ))}

          {/* Right Arrow - visible when more than one page */}
          {this.hasMultiplePages() && (
            <button
              className={`${styles.arrows} ${styles.arrowRight}`}
              onClick={this.handleNextMeetingPage}
              data-action="tabAction"
              onKeyDown={(event) => { if (event.key == "Enter") this.handleNextMeetingPage }}
            />
          )}

        </div>

        {/* Pagination rendering logic - dots */}
        {this.hasMultiplePages() && (
          <div className={styles.paginationWrapper}>
            <div className={styles.paginationCircles}>
              {Array.from({
                length: this.getTotalPages(),
              }, (_, index) => (
                <span
                  key={index}
                  className={styles.paginationIconWrapper}
                  onClick={() => this.setCurrentPage(index + 1)}
                  onKeyDown={(event) => { if (event.key == "Enter") this.setCurrentPage(index + 1) }}

                >
                  <Icon
                    iconName={
                      this.state.currentMeetingPage === index + 1
                        ? "RadioBtnOn" // Show the filled icon for the selected page
                        : "CircleRing"  // Show the default outline icon for unselected pages
                    }
                    className={`${styles.paginationIcon}}`}
                    data-action="tabAction"
                  />
                </span>
              ))}
            </div>
          </div>
        )}


        {/* Dialog - show meetings */}
        {this.state.isDialogVisible && this.state.selectedMeeting && (
          <div>
            <MeetingManagementDetails
              isDialogVisible={this.state.isDialogVisible}
              selectedMeeting={this.state.selectedMeeting}
              showRegistrations={this.state.showRegistrations}
              registrationsCount={this.state.registrationsCount}
              registeredUsers={this.state.registeredUsers}
              isUserRegistered={this.state.isUserRegistered}
              toggleRegistrations={this.props.toggleRegistrations}
              registerOptions={this.props.registerOptions}
              currentDate={this.currentTime}
              onRegister={() => this.registerForMeeting(this.state.selectedMeeting)}
              onUnregister={() => this.unregisterFromMeeting(this.state.selectedMeetingId)}
              closeDialog={this.closeDialog}
              downloadICalFile={() => this.downloadICalFile(this.state.selectedMeeting)}
              toggleShowRegistrations={() =>
                this.setState((prevState) => ({ showRegistrations: !prevState.showRegistrations }))
              }
            />
          </div>

        )}


        {/* Confirmation Dialog */}
        {this.state.confirmationDialogVisible &&
          <MeetingManagementConfirmation
            isVisible={this.state.confirmationDialogVisible}
            message={this.state.confirmationDialogMessage}
            buttonText={this.state.confirmationDialogButtonText}
            title={this.state.confirmationDialogTitle}
            onConfirm={() => {
              if (this.state.confirmationDialogAction) {
                this.state.confirmationDialogAction(); // Udfører handling
              }
              this.closeConfirmationDialog(); // Lukker dialogen
            }}
            onClose={this.closeConfirmationDialog}
            requiresApproval={this.state.requiresApproval}
            approvalMessage={this.state.approvalMessage}
          />
        }

        {this.state.successDialogVisible &&
          <MeetingManagementSuccess
            isVisible={this.state.successDialogVisible}
            message={this.state.successDialogMessage}
            buttonText={this.state.successDialogButtonText}
            title={this.state.successDialogTitle}
            onClose={this.closeSuccessDialog}
          />
        }

      </div>
    );
  }


  public async componentDidMount(): Promise<void> {
    try {
      // First, fetch the current user data
      await this.getCurrentUser();

      const [registeredMeetingIds, meetings, allRegistrations] = await Promise.all([
        this.getUserRegistrations(),
        this.getMeetings(),
        this.getAllRegistrations(), // Fetch all registrations here
      ]);


      const meetingsCount = meetings.length;

      // Calculate registration counts for each meeting from `allRegistrations`
      const registrationCounts = allRegistrations.reduce((counts, reg) => {
        counts[reg.MeetingId] = (counts[reg.MeetingId] || 0) + 1;
        return counts;
      }, {} as { [meetingId: number]: number });

      this.determineLayout(this.props._width, meetingsCount);

      this.setState({
        //waitingListMeetingIds: waitingListMeetingIds,
        registeredMeetingIds: registeredMeetingIds,
        meetings: meetings,
        registrationCounts,
        meetingsCount: meetingsCount,
        dataLoaded: true,
        allRegistrations, // Store all registrations in state
      });

      this.checkUserCanManageLists();
    } catch (error) {
      console.error("Error loading data:", error);
      this.setState({ dataLoaded: true });
    }
  }


  private async getCurrentUser(): Promise<void> {
    try {
      // Fetch current user information
      const currentUser = await this.props.sp.web.currentUser();
      const userId = currentUser.Id;
      const userName = currentUser.Title;
      const userEmail = currentUser.Email

      // Set user data in state
      this.setState({
        userId,
        userName,
        userEmail,
      });

      console.log(`User data set: ID = ${userId}, Name = ${userName}`);
    } catch (error) {
      console.error("Error fetching user information:", error);
    }
  }




  private async getRegistrationList() {
    if (!this.registrationList) {
      this.registrationList = this.props.sp.web.lists.getById(this.props.meetingRegistrationListOption);
    }
    return this.registrationList;
  }

  private async getAllRegistrations(): Promise<any[]> {
    try {
      const registrationList = await this.getRegistrationList();
      const allRegistrations = await registrationList.items.top(5000)(); // Retrieve all registrations

      return allRegistrations;
    } catch (error) {
      console.error("Error fetching all registrations:", error);
      return [];
    }
  }

  private async getMeetingRegistrations(meetingId: number): Promise<void> {
    try {
      // Filter the registrations from allRegistrations in the state based on the meetingId
      const registrations = this.state.allRegistrations.filter((reg) => reg.MeetingId === meetingId);

      // Map through filtered registrations and fetch user names and pictures
      const registeredUsers = await Promise.all(
        registrations.map(async (reg) => {
          const userName = reg.Title.split(':')[1].trim();
          const userEmail = reg.Email || ''; // assuming `UserEmail` is available
          const pictureUrl = await this.getUserProfilePicture(userEmail);
          return { userName, pictureUrl };
        })
      );


      // Update the component's state with the count and users for this meeting
      this.setState({
        registrationsCount: registrations.length,
        registeredUsers,
      });
    } catch (error) {
      console.error("Error loading registrations count:", error);
    }
  }



  private async getUserRegistrations(): Promise<number[]> {
    try {
      const registrationList = await this.getRegistrationList();


      const userRegistrations = await registrationList.items
        .filter(`UserId eq ${this.state.userId}`)
        .top(5000)();

      return userRegistrations.map((registration) => registration.MeetingId);
    } catch (error) {
      console.error("Error loading registrations:", error);
      return [];
    }
  }



  private async getUserProfilePicture(userEmail: string): Promise<string | null> {
    try {
      const graphClient = await this.props.context.msGraphClientFactory.getClient("3");
      const response = await graphClient
        .api(`/users/${userEmail}/photo/$value`)
        .responseType("blob")
        .get();

      const url = window.URL.createObjectURL(response);
      console.log("Profile picture URL:", url);
      return url;
    } catch (error) {
      console.error(`Error fetching profile picture for ${userEmail}:`, error);
      return `/_layouts/15/userphoto.aspx?size=L&username=${userEmail}`; // fallback URL
    }
  }




  private getTotalPages(): number {
    const filteredMeetings = this.filterMeetings();
    const totalMeetingsToShow = Math.min(filteredMeetings.length, this.props.numberOfMeetings || 4);
    return Math.ceil(totalMeetingsToShow / this.state.meetingsPerPage);
  }



  private hasMultiplePages(): boolean {
    return this.getTotalPages() > 1;
  }


  private handleNextMeetingPage = () => {
    const totalPages = this.getTotalPages();

    this.setState((prevState) => ({
      currentMeetingPage: prevState.currentMeetingPage >= totalPages
        ? 1
        : prevState.currentMeetingPage + 1
    }));
  };



  private handlePreviousMeetingPage = () => {
    const totalPages = this.getTotalPages();

    this.setState((prevState) => ({
      currentMeetingPage: prevState.currentMeetingPage <= 1
        ? totalPages
        : prevState.currentMeetingPage - 1
    }));
  };


  private handleToggleMyRegistrations = (checked: boolean) => {
    this.setState({ showMyRegistrations: checked }, () => {
      const totalPages = this.getTotalPages();

      if (this.state.currentMeetingPage > totalPages) {
        this.setState({ currentMeetingPage: 1 });
      }
    });
  };





  private getPaginatedMeetings(filteredMeetings: IMeeting[]): IMeeting[] {
    const totalMeetingsToShow = Math.min(filteredMeetings.length, this.props.numberOfMeetings || 4);
    const meetingsToDisplay = filteredMeetings.slice(0, totalMeetingsToShow);

    // Fejl: Beregningen af startIndex bruger `currentMeetingPage` i stedet for `meetingsPerPage`
    const startIndex = (this.state.currentMeetingPage - 1) * this.state.meetingsPerPage;
    const endIndex = Math.min(startIndex + this.state.meetingsPerPage, meetingsToDisplay.length);

    return meetingsToDisplay.slice(startIndex, endIndex);
  }


  private getJustifyClass(itemsOnCurrentPage: number): string {
    if (this.state.meetingsPerPage > 2) {
      if ((this.state.layout === 'mediumLayout' && itemsOnCurrentPage === 2)) {
        return `${styles.mediumLayoutAdjustment} ${styles.cellBoxPadding}`;
      } else if ((this.state.layout === 'largeLayout' && (itemsOnCurrentPage === 2 || itemsOnCurrentPage === 3))) {
        return `${styles.largeLayoutAdjustment} ${styles.cellBoxPadding}`;
      }
      return '';
    }
    return '';
  }

  private handleSeeAllClick = () => {
    window.open(this.props.seeAllUrl, '_blank');
  };

  private setCurrentPage = (pageNumber: number) => {
    this.setState({ currentMeetingPage: pageNumber });
  };


  private determineLayout(width: number, meetingsCount: number) {
    let layout = this.state.layout;
    let meetingsPerPage = this.state.meetingsPerPage;

    // Adjust layout and meetings per page based on width
    if (width >= 1072) {
      meetingsPerPage = 4;
      layout = 'largeLayout';
      console.log(width + ' - Stort layout')
    } else if (width >= 738.667) {
      meetingsPerPage = 3;
      layout = 'mediumLayout';
      console.log(width + ' - Mellem layout')
    } else if (width >= 558.667) {
      meetingsPerPage = 2;
      layout = 'largeLayout';
      console.log(width + ' - Mellem-lille layout')
    } else {
      meetingsPerPage = 10;
      layout = 'smallLayout';
      console.log(width + ' - Lille layout')
    }

    this.setState({ layout, meetingsPerPage: meetingsPerPage }, () => {
      if ((layout === 'largeLayout' && meetingsCount > 4) ||
        (layout === 'mediumLayout' && meetingsCount > 3)) {
        this.setState({ showPagination: true });
      } else {
        this.setState({ showPagination: false });
      }

    });
  }



  private generateICalContent(meeting: IMeeting, isCancelled: boolean = false, attendees: string[] = []): string {
    const startTime = new Date(meeting.StartTime).toISOString().replace(/-|:|\.\d\d\d/g, '');
    const endTime = new Date(meeting.EndTime).toISOString().replace(/-|:|\.\d\d\d/g, '');
    const room = meeting.Room.replace(/,/g, '\\,').replace(/\n/g, '\\n');
    const description = meeting.Description.replace(/,/g, '\\,').replace(/\n/g, '\\n');

    const attendeeList = attendees.map(email => `ATTENDEE;RSVP=TRUE:mailto:${email}`).join('\n');
    const status = isCancelled ? "CANCELLED" : "CONFIRMED";

    return `
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Your Company//Meeting Management//EN
BEGIN:VEVENT
UID:${meeting.Id}@yourdomain.com
DTSTAMP:${new Date().toISOString().replace(/-|:|\.\d\d\d/g, '')}
DTSTART:${startTime}
DTEND:${endTime}
SUMMARY:${meeting.Title}
ROOM:${room}
DESCRIPTION:${description}
STATUS:${status}
${attendeeList}
END:VEVENT
END:VCALENDAR
`.trim();
  }

  private downloadICalFile = (meeting: IMeeting, isCancelled: boolean = false, attendees: string[] = []) => {
    const icalContent = this.generateICalContent(meeting, isCancelled, attendees);
    const blob = new Blob([icalContent], { type: 'text/calendar' });
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `${meeting.Title}.ics`;
    a.click();

    window.URL.revokeObjectURL(url);
  };




  private async getMeetings(): Promise<IMeeting[]> {
    try {
      // Fetch list items directly from SharePoint
      const listItems = await this.props.sp.web.lists.getById(this.props.meetingsListOption).items.select(
        "Id",
        "UniqueId",
        "Title",
        "StartTime",
        "EndTime",
        "RegistrationDeadline",
        "Room",
        "Description",
        "MeetingType",
        "Approval",
        "Category",
        "MaxRegistrations"
      )();

      // Filter out meetings that have already started
      const upcomingMeetings = listItems.filter(item => new Date(item.StartTime) > this.currentTime);

      // Map list items to the IMeeting structure
      const mappedMeetings: IMeeting[] = upcomingMeetings.map(meeting => ({
        Id: meeting.Id,
        Title: meeting.Title,
        StartTime: meeting.StartTime,
        EndTime: meeting.EndTime,
        RegistrationDeadline: meeting.RegistrationDeadline ? meeting.RegistrationDeadline : meeting.StartTime,
        Description: meeting.Description,
        MeetingType: meeting.MeetingType,
        Room: meeting.Room,
        Category: meeting.Category,
        Approval: meeting.Approval,
        MaxRegistrations: meeting.MaxRegistrations,
      }));

      // Sort the meetings by StartTime
      const sortedMeetings = mappedMeetings.sort((a, b) => new Date(a.StartTime).getTime() - new Date(b.StartTime).getTime());

      // Return the sorted meetings
      return sortedMeetings;

    } catch (error) {
      console.error("Error loading meetings:", error);
      return [];
    }
  }





  private filterMeetings(): IMeeting[] {

    let { meetings, selectedTypeFilter, selectedCategoryFilter, selectedRoomFilter, showMyRegistrations, registeredMeetingIds } = this.state;
    const { filterByMeetingType, filterByCategory, filterByRoom, filterById } = this.props;

    const today = new Date();

    // Filter meetings based on EndDate and other conditions
    let filteredMeetings = meetings.filter(meeting => {
      const endTime = new Date(meeting.EndTime);

      // Include meeting only if it has no end date or the end date is today or in the future
      return !meeting.EndTime || endTime >= today;
    });

    // Apply filterById if it is provided
    if (filterById) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.Id === filterById);
    }

    // Apply the editor-selected filters (from properties)
    if (filterByMeetingType && filterByMeetingType.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.MeetingType === filterByMeetingType);
    }

    if (filterByCategory && filterByCategory.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.Category === filterByCategory);
    }

    if (filterByRoom && filterByRoom.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.Room === filterByRoom);
    }

    // Apply the user-selected filters (from state)
    if (selectedTypeFilter && selectedTypeFilter.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.MeetingType === selectedTypeFilter);
    }

    if (selectedCategoryFilter && selectedCategoryFilter.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.Category === selectedCategoryFilter);
    }

    if (selectedRoomFilter && selectedRoomFilter.length > 0) {
      filteredMeetings = filteredMeetings.filter(meeting => meeting.Room === selectedRoomFilter);
    }


    // If the "Show My Registrations" toggle is active, filter by registrations
    if (showMyRegistrations) {
      filteredMeetings = filteredMeetings.filter(meeting => registeredMeetingIds.includes(meeting.Id));
    }

    return filteredMeetings;
  }






  private showMeetingDetails = async (meetingId: number): Promise<void> => {

    const test = this.currentTime;
    const currentMeetings: IMeeting[] = this.state.meetings;

    const selectedMeetings = currentMeetings.filter(meeting => meeting.Id === meetingId);
    if (selectedMeetings && selectedMeetings.length > 0) {
      const selectedMeeting = selectedMeetings[0];
      const isUserRegistered = this.state.registeredMeetingIds.includes(meetingId);
      //const isUserOnWaitingList = await this.checkUserOnWaitingList(meetingId);


      this.setState({
        selectedMeetingId: meetingId,
        isDialogVisible: true,
        selectedMeeting: selectedMeeting,
        isUserRegistered,
        //isUserOnWaitingList
      }, () => {
        this.getMeetingRegistrations(meetingId);
      });
    }
  }

  private closeDialog = () => {
    this.setState({ isDialogVisible: false, selectedMeeting: null, registrationsCount: 1, showRegistrations: false });
  }


  private async registerForMeeting(meeting: IMeeting): Promise<void> {
    try {
      const message = strings.MessageConfirmMeeting;
      const approvalMessage = meeting.Approval ? strings.ApprovalMessage : "";
      const buttonText = strings.Register;
      const title = strings.Confirmation;

      this.showConfirmationDialog(
        message,
        buttonText,
        title,
        meeting.Approval,
        approvalMessage,
        async () => {
          await this.closeConfirmationDialog();

          try {
            const registrationList = this.props.sp.web.lists.getById(this.props.meetingRegistrationListOption);

            const itemToAdd = {
              Title: `Username: ${this.state.userName}`,
              UserId: this.state.userId,
              MeetingId: meeting.Id,
              MeetingName: meeting.Title,
              Email: this.state.userEmail
            };


            await registrationList.items.add(itemToAdd);
            this.showSuccessDialog(strings.Registered);

            this.setState(prevState => ({
              registeredMeetingIds: [...prevState.registeredMeetingIds, meeting.Id]
            }));

          } catch (error) {
            console.error("Error registering for meeting:", error);
          }
        }
      );
    } catch (error) {
      console.error("Error during registration:", error);
    }
  }



  private showConfirmationDialog = (
    message: string,
    buttonText: string,
    title: string,
    requiresApproval: boolean = false,
    approvalMessage: string = '',
    onConfirm: () => void,
  ) => {
    this.setState({
      confirmationDialogVisible: true,
      confirmationDialogMessage: message,
      confirmationDialogButtonText: buttonText,
      confirmationDialogTitle: title,
      confirmationDialogAction: onConfirm,
      requiresApproval,
      approvalMessage,
    });
  };


  private showSuccessDialog = (message: string) => {
    this.setState({
      successDialogVisible: true,
      successDialogMessage: message,

    });
  };

  private closeConfirmationDialog = () => {
    this.setState({
      confirmationDialogVisible: false,
      confirmationDialogMessage: '',
      confirmationDialogAction: null,
      requiresApproval: false,
      approvalMessage: '',
    });
  };

  private closeSuccessDialog = () => {
    this.setState({
      successDialogVisible: false,
      successDialogMessage: ''
    }, () => window.location.reload());
  };




  private async unregisterFromMeeting(meetingId: number): Promise<void> {
    const message = strings.MessageUnregisterMeeting;
    const buttonText = strings.Unregister;
    const title = strings.Confirmation;

    this.showConfirmationDialog(
      message,
      buttonText,
      title,
      false, // No approval required
      "",
      async () => {
        this.closeConfirmationDialog();
        try {
          const registrationList = this.props.sp.web.lists.getById(this.props.meetingRegistrationListOption);

          const items: any[] = await registrationList.items
            .filter(`MeetingId eq ${meetingId} and UserId eq ${this.state.userId}`)
            .top(1)();

          if (items.length > 0) {
            const itemId = items[0].Id;
            await registrationList.items.getById(itemId).delete();
            this.showSuccessDialog(strings.Unregistered);

            this.setState(prevState => ({
              registeredMeetingIds: prevState.registeredMeetingIds.filter(id => id !== meetingId)
            }));
          }
        } catch (error) {
          console.error("Error unregistering from meeting:", error);
        }
      }
    );
  }



  private renderShimmer() {
    return (

      <div>
        <Shimmer className={styles.shimmerHeader} width="20%"
          shimmerElements={[
            { type: ShimmerElementType.gap, height: 23 },
          ]}
        />
        <Shimmer className={styles.shimmerText} width="45%"
          shimmerElements={[
            { type: ShimmerElementType.gap, height: 10 },
          ]}
        />
        <Shimmer className={styles.shimmerFilter} width="45%"
          shimmerElements={[
            { type: ShimmerElementType.gap, height: 35 },
          ]}
        />
        <Shimmer className={styles.shimmerButton} width="15%"
          shimmerElements={[
            { type: ShimmerElementType.gap, height: 17 },
          ]}
        />

        <div className={styles.shimmerContainer}>

          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 60, width: 60 },
            ]}

          />

          <div className={styles.shimmerDetails}>
            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />
          </div>

        </div>
        <div className={styles.shimmerContainer}>

          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 60, width: 60 },
            ]}

          />

          <div className={styles.shimmerDetails}>
            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />
          </div>

        </div>
        <div className={styles.shimmerContainer}>

          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 60, width: 60 },
            ]}

          />

          <div className={styles.shimmerDetails}>
            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />

            <Shimmer className={styles.shimmerLines} width="85%"
              shimmerElements={[
                { type: ShimmerElementType.gap, height: 10 },
              ]}
            />
          </div>

        </div>
      </div>



    );
  }

  private async checkUserCanManageLists(): Promise<void> {
    try {

      const currentPermission: SPPermission = this.props.context.pageContext.web.permissions;
      const isAbleToManage: boolean = currentPermission.hasPermission(SPPermission.manageLists) && currentPermission.hasPermission(SPPermission.managePermissions);

      console.log(`Current user permission: { High: ${currentPermission.value.High}, Low: ${currentPermission.value.Low} }`);
      console.log(`Current user is${isAbleToManage ? ' ' : ' not '}able to manage lists and permissions.`);


      this.setState({ isEditor: isAbleToManage });
    } catch (error) {
      console.error("Error checking user permissions:", error);
    }
  }
}