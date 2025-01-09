import * as React from 'react';
import styles from './MeetingManagement.module.scss'; // Adjust the import path as necessary
import { IMeeting } from './IMeeting';
import * as strings from 'MeetingManagementWebPartStrings'
import { PrimaryButton } from 'office-ui-fabric-react';
import { DefaultButton } from '@fluentui/react';


interface IMeetingManagementCellProps {
    key: number;
    meeting: IMeeting;
    layout: string;
    registrationCounts: { [meetingId: number]: number };
    registeredMeetingIds: number[];
    currentDate: Date;
    showMeetingDetails: (meetingId: number) => void;
    registerForMeeting: (meeting: IMeeting) => void;
    unregisterFromMeeting: (meetingId: number) => void;
    justifyClass: string;
    registerOptions: string;
    responsibleUser: { userName: string; userEmail: string };
    getUserProfilePicture: (userEmail: string) => Promise<string | null>;
}

const MeetingManagementCell: React.FC<IMeetingManagementCellProps> = ({
    meeting,
    layout,
    registrationCounts,
    registeredMeetingIds,
    currentDate,
    showMeetingDetails,
    registerForMeeting,
    unregisterFromMeeting,
    justifyClass,
    registerOptions,
    responsibleUser,
    getUserProfilePicture,
}) => {

    const [profilePictureUrl, setProfilePictureUrl] = React.useState<string | null>(null);

    React.useEffect(() => {
        const fetchProfilePicture = async () => {
            if (responsibleUser?.userEmail) {
                const url = await getUserProfilePicture(responsibleUser.userEmail);
                setProfilePictureUrl(url);
            }
        };
        fetchProfilePicture();
    }, [responsibleUser, getUserProfilePicture]);

    const startDate = new Date(meeting.StartTime);
    const registrationDeadline = new Date(meeting.RegistrationDeadline);
    const isRegistrationClosed = currentDate > registrationDeadline;

    const registrationCount = registrationCounts[meeting.Id] || 0;

    const day = startDate.getDate();
    const month = startDate.toLocaleString("da-DK", { month: 'short' }).toUpperCase();
    const date = startDate.toLocaleDateString("da-DK", { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' });
    const time = startDate.toLocaleTimeString("da-DK", { hour: '2-digit', minute: '2-digit' });

    const isRegistered = registeredMeetingIds.includes(meeting.Id);
    const isEventFull = registrationCount >= meeting.MaxRegistrations;

    let imageUrl = '';
    let cellDateStyle = {};
    let cellDateStyleSmall = {};
    let smallLayoutMonthStyle = {};
    let smallLayoutDayStyle = {};

    /* if (meeting.Background) {
        try {

            imageUrl = meeting.Background ? meeting.Background : "";

            cellDateStyle = {
                color: '#ffffff',
                background: 'linear-gradient(to top, rgb(0 0 0 / 54%), rgba(255, 255, 255, 0))',
            };

            cellDateStyleSmall = {
                padding: '14px 0 0 4px',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'flex-start',
                backgroundColor: 'hsla(0, 0%, 0%, 0.29)',
                borderRadius: '1px',
            }

            smallLayoutMonthStyle = {
                fontSize: '12px',
                fontWeight: 600,
                color: 'white',
            };

            smallLayoutDayStyle = {
                fontSize: '24px',
                fontWeight: 600,
                marginTop: '-2px',
                color: 'white',
            };

        } catch (error) {
            console.error('Error parsing Background field:', error);
        }
    }
 */

    const backgroundStyle = {
        backgroundImage: `url("${imageUrl}")`,
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        width: '100%',
    };

    return (
        <div className={`${justifyClass}`}>
            {(layout === 'largeLayout' || layout === 'mediumLayout') && (
                <div
                    className={`${styles.cellBox} ${styles[layout]}`}
                    onClick={() => showMeetingDetails(meeting.Id)}
                >

                    <div
                        onClick={() => showMeetingDetails(meeting.Id)}
                        onKeyDown={(event) => { if (event.key === "Enter") showMeetingDetails(meeting.Id) }}
                        data-action="tabAction"
                    >
                        <div className={`${styles.cellBackground} ${styles[layout]}`} style={backgroundStyle}>
                            <div className={`${styles.cellHeadline} ${styles[layout]}`} style={backgroundStyle}>
                                <div className={`${styles.cellTitle} ${styles[layout]}`}>{meeting.Title}</div>
                                {/* <div className={`${styles.cellDate} ${styles[layout]}`} style={cellDateStyle}>
                                <div className={`${styles.month} ${styles[layout]}`}>{month?.toLowerCase()} </div>
                                <div className={`${styles.day} ${styles[layout]}`}>{day}</div>
                            </div> */}

                                <div className={styles.room}>{meeting.Room}</div>
                            </div>

                            <div className={`${styles.cellSubTitle} ${styles[layout]}`}>{date}, {time}</div>

                        </div>

                        <div className={styles.line}></div>

                        <div className={`${styles.cellContent} ${styles[layout]} `}>

                            <div className={styles.category}>{meeting.Category}</div>

                            <div className={`${styles.responsibleUserTitle} ${styles[layout]}`}>Mødeansvarlig </div>
                            <div className={styles.responsibleUser}>
                                {profilePictureUrl && (
                                    <img
                                        src={profilePictureUrl}
                                        alt={`${responsibleUser.userName}'s profile`}
                                        className={styles.profilePicture}
                                    />
                                )}
                                {responsibleUser.userName}
                            </div>



                            <div className={`${styles.cellAgendaTitle} ${styles[layout]}`}>Agenda</div>
                            <div className={`${styles.cellAgenda} ${styles[layout]}`}>{meeting.Description}</div>



                        </div>

                        <div className={styles.line}></div>
                    </div>



                    {(registerOptions === 'showInView' || registerOptions === 'showBoth') && (
                        <div className={styles.buttonSektion}>
                            {!isRegistered && !isEventFull && !isRegistrationClosed && (
                                <PrimaryButton
                                    //className={`${styles.cellRegistrationButton} ${styles[layout]}`}
                                    onClick={(e) => { e.stopPropagation(); registerForMeeting(meeting); }}
                                    data-action="tabAction"
                                >
                                    {"Deltag"}
                                </PrimaryButton>
                            )}
                            {isRegistered && (
                                <DefaultButton
                                    //className={`${styles.cellRegistrationButton} ${styles[layout]}`}
                                    onClick={(e) => { e.stopPropagation(); unregisterFromMeeting(meeting.Id); }}
                                    data-action="tabAction"
                                >
                                    {"Deltag ikke"}
                                </DefaultButton>
                            )}
                         
                        </div>
                    )}

                </div>
            )}

            {layout === 'smallLayout' && (
                <div
                    className={`${styles.cellBox} ${styles[layout]}`}
                    onClick={() => showMeetingDetails(meeting.Id)}
                >

                    <div
                        onClick={() => showMeetingDetails(meeting.Id)}
                        onKeyDown={(event) => { if (event.key === "Enter") showMeetingDetails(meeting.Id) }}
                        data-action="tabAction"
                    >
                        <div className={styles.cell}>
                            <div className={`${styles.cellBackground} ${styles[layout]}`} style={backgroundStyle}>
                                <div className={`${styles.cellDate} ${styles[layout]}`} style={cellDateStyleSmall}>
                                    <div className={`${styles.month} ${styles[layout]}`} style={smallLayoutMonthStyle}>{month}</div>
                                    <div className={`${styles.day} ${styles[layout]}`} style={smallLayoutDayStyle}>{day}</div>
                                </div>
                            </div>
                        </div>
                    </div>


                    <div className={`${styles.cellContent} ${styles[layout]} `}>
                        <div className={`${styles.cellTitle} ${styles[layout]}`}>{meeting.Title}</div>
                        <div className={`${styles.cellSubTitle} ${styles[layout]}`}>{date}, {time}</div>
                        <div className={styles.buttonSektion}>
                            {(registerOptions === 'showInView' || registerOptions === 'showBoth') && (
                                <div>
                                    {!isRegistered && !isEventFull && !isRegistrationClosed && (
                                        <button
                                            className={`${styles.cellRegistrationButton} ${styles[layout]}`}
                                            onClick={(e) => { e.stopPropagation(); registerForMeeting(meeting); }}
                                            data-action="tabAction"
                                        >
                                            {strings.Register}
                                        </button>
                                    )}
                                    {isRegistered && (
                                        <button
                                            className={`${styles.cellRegistrationButton} ${styles[layout]}`}
                                            onClick={(e) => { e.stopPropagation(); unregisterFromMeeting(meeting.Id); }}
                                            data-action="tabAction"
                                        >
                                            {strings.Unregister}
                                        </button>
                                    )}
                               
                                </div>
                            )}
                        </div>
                    </div>

                </div>
            )}
        </div>
    );
};

export default MeetingManagementCell;

