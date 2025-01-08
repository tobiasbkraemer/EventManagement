import * as React from 'react';
import { Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton, IconButton, Persona, PersonaSize, PersonaPresence } from '@fluentui/react';
import styles from './MeetingManagement.module.scss';
import * as strings from 'MeetingManagementWebPartStrings'


interface IMeetingManagementDetailsProps {
    isDialogVisible: boolean;
    selectedMeeting: any; 
    showRegistrations: boolean;
    currentDate: Date;
    registrationsCount: number;
    registeredUsers: { userName: string, pictureUrl: string }[];
    isUserRegistered: boolean;
    toggleRegistrations: boolean;
    registerOptions: string;
    onRegister: () => void;
    onUnregister: () => void;
    closeDialog: () => void;
    downloadICalFile: () => void;
    toggleShowRegistrations: () => void;
}

const MeetingManagementDetails: React.FC<IMeetingManagementDetailsProps> = ({
    isDialogVisible,
    selectedMeeting: selectedMeeting,
    showRegistrations,
    currentDate,
    registrationsCount,
    registeredUsers,
    isUserRegistered,
    toggleRegistrations,
    registerOptions,
    onRegister,
    onUnregister,
    closeDialog,
    downloadICalFile,
    toggleShowRegistrations,
}) => {
    if (!isDialogVisible || !selectedMeeting) return null;


    return (
        <Dialog
            hidden={!isDialogVisible}
            onDismiss={closeDialog}
            dialogContentProps={{
                type: DialogType.normal,
                closeButtonAriaLabel: 'Close',
                titleProps: {
                    style: { padding: '16px 24px 0px' },
                },
                title:
                    (
                        <div className={styles.dialogHeader}>
                            <div className={styles.dialogTitle}>
                                <span className={styles.dialogTitle}>{selectedMeeting ? selectedMeeting.Title : ''}</span>
                                
                                <div>
                                    <IconButton
                                        styles={{
                                            root: { height: 'auto' },
                                            rootHovered: { backgroundColor: 'transparent' },
                                            rootPressed: { backgroundColor: 'transparent' },
                                            icon: { color: '#000' }
                                        }}
                                        iconProps={{ iconName: 'Cancel' }}
                                        ariaLabel="Close"
                                        onClick={closeDialog}
                                    />

                                </div>
                            </div>

                        </div>
                    )
            }}
            containerClassName={styles.dialogContainer}
        >
            <div className={styles.dialogBox}>

                <div className={styles.dialogDeadlineSektion}>
                    <div className={styles.dialogRegistrationDeadline}>
                        <b>{strings.RegistrationDeadline}:</b>&nbsp;
                        {new Date(selectedMeeting.RegistrationDeadline).toLocaleDateString("da-DK", {
                            weekday: 'long',
                            year: 'numeric',
                            month: 'long',
                            day: 'numeric'
                        }).replace(/^\w/, (c) => c.toUpperCase())}
                        , {new Date(selectedMeeting.RegistrationDeadline).toLocaleTimeString("da-DK", {
                            hour: '2-digit',
                            minute: '2-digit'
                        })}
                    </div>
                </div>


                <div className={styles.dialogDescription}>
                    <p>{selectedMeeting.Description}</p>
                </div>
                <div className={styles.dialogBody}>
                    <div className={styles.dialogDetails}>
                        <div className={styles.dialogDetailItem}>
                            <h3>{strings.DateAndTime}</h3>
                            <p>
                                {new Date(selectedMeeting.StartTime).toLocaleDateString("da-DK", {
                                    weekday: 'long',
                                    year: 'numeric',
                                    month: 'long',
                                    day: 'numeric'
                                }).replace(/^\w/, (c) => c.toUpperCase())}
                                , {new Date(selectedMeeting.StartTime).toLocaleTimeString("da-DK", {
                                    hour: '2-digit',
                                    minute: '2-digit'
                                })}
                            </p>
                        </div>
                        <div className={styles.dialogDetailItem}>
                            <h3>{strings.Room}</h3>
                            <p>{selectedMeeting.Room}</p>
                        </div>
                        <div className={styles.dialogDetailItem}>
                            <h3>{strings.Category}</h3>
                            <p>{selectedMeeting.Category}</p>
                        </div>
                        <div className={styles.dialogDetailItem}>
                            <h3>{strings.EndDateAndTime}</h3>
                            <p>
                                {new Date(selectedMeeting.EndTime).toLocaleDateString("da-DK", {
                                    weekday: 'long',
                                    year: 'numeric',
                                    month: 'long',
                                    day: 'numeric'
                                }).replace(/^\w/, (c) => c.toUpperCase())}
                                , {new Date(selectedMeeting.EndTime).toLocaleTimeString("da-DK", {
                                    hour: '2-digit',
                                    minute: '2-digit'
                                })}
                            </p>
                        </div>
                        <div className={styles.dialogDetailItem}>
                            <h3>{strings.MeetingType}</h3>
                            <p>{selectedMeeting.MeetingType}</p>
                        </div>
                        {toggleRegistrations && (
                            <div className={styles.dialogDetailItem}>
                                <h3>{strings.Registrations}</h3>
                                <p>
                                    {registrationsCount}/{selectedMeeting.MaxRegistrations}

                                </p>
                            </div>
                        )}
                    </div>
                </div>

                <DialogFooter
                    styles={{
                        actionsRight: {
                            display: 'flex',
                            justifyContent: 'space-between',
                            width: '100%',
                            //flexDirection: 'row-reverse'
                        },
                    }}>
                    <div className={styles.dialogButtons}>

                        <DefaultButton
                            text={strings.Close}
                            onClick={closeDialog}
                        />

                        <div className={styles.dialogFooterButton}>
                            <PrimaryButton text={strings.DownloadICal} onClick={downloadICalFile} />
                        </div>

                        <PrimaryButton
                            text={showRegistrations ? "Vis ikke deltagere" : "Vis deltagere"}
                            onClick={toggleShowRegistrations}
                            disabled={!toggleRegistrations || registrationsCount <= 0}
                        />






                    </div>


                    {(registerOptions === 'showInEvent' || registerOptions === 'showBoth') && (
                        <div>
                            {!isUserRegistered
                                && registrationsCount < selectedMeeting.MaxRegistrations
                                && (selectedMeeting.RegistrationDeadline === null || new Date(selectedMeeting.RegistrationDeadline) > currentDate)
                                && (
                                    <PrimaryButton text={"Deltag"} onClick={onRegister} />
                                )
                            }

                            {isUserRegistered && (
                                <DefaultButton text={"Deltag ikke"} onClick={onUnregister} />
                            )}

                        </div>

                    )}
                </DialogFooter>

                {showRegistrations && (
                    <div>


                        <div className={styles.scrollableContainer}>
                            {registeredUsers.map((user, index) => (
                                <div className={styles.registration} key={index}>
                                    <Persona
                                        text={user.userName}
                                        imageUrl={user.pictureUrl || undefined}
                                        imageInitials={user.userName.split(' ').map(name => name[0]).join('')}
                                        size={PersonaSize.size24}
                                        presence={PersonaPresence.online}
                                        hidePersonaDetails={false}
                                        imageAlt={user.userName}
                                    />
                                </div>
                            ))}
                        </div>
                    </div>
                )}
            </div>



        </Dialog>
    );
};

export default MeetingManagementDetails;
