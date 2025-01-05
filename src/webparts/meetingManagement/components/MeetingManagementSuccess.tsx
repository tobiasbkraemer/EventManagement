import * as React from 'react';
import { Dialog, DialogFooter, PrimaryButton, DialogType } from '@fluentui/react';
import * as strings from 'MeetingManagementWebPartStrings';

interface IMeetingManagementSuccessProps {
    message: string;
    buttonText: string;
    title: string;
    isVisible: boolean;
    onClose: () => void;

}

const MeetingManagementSuccess: React.FC<IMeetingManagementSuccessProps> = ({
    message,
    isVisible,
    onClose,
}) => {

    return (
        <Dialog
            hidden={!isVisible}
            onDismiss={onClose}
            dialogContentProps={{
                type: DialogType.normal,
                title: strings.Success,
                subText: message,
            }}
            modalProps={{
                isBlocking: true,
            }}
        >
            <DialogFooter>
                <PrimaryButton text={strings.OK} onClick={onClose} />
            </DialogFooter>
        </Dialog>
    );
};

export default MeetingManagementSuccess;
