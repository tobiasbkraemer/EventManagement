import * as React from 'react';
import { Dialog, DialogFooter, PrimaryButton, DefaultButton, DialogType, Checkbox } from '@fluentui/react';
import * as strings from 'MeetingManagementWebPartStrings'


interface IMeetingManagementConfirmationProps {
    message: string;
    isVisible: boolean;
    buttonText: string;
    title: string;
    onConfirm: () => void;
    onClose: () => void;
    requiresApproval?: boolean;
    approvalMessage?: string;
}

const MeetingManagementConfirmation: React.FC<IMeetingManagementConfirmationProps> = ({
    message,
    buttonText,
    title,
    isVisible,
    onConfirm,
    onClose,
    requiresApproval = false,
    approvalMessage = '',
}) => {
    const [isApproved, setIsApproved] = React.useState(!requiresApproval);

    // Reset isApproved state based on requiresApproval whenever dialog visibility changes
    React.useEffect(() => {
        setIsApproved(!requiresApproval);
    }, [isVisible, requiresApproval]);

    const handleApprovalChange = (_, checked?: boolean) => {
        setIsApproved(!!checked);
    };

    return (
        <Dialog
            hidden={!isVisible}
            onDismiss={onClose}
            dialogContentProps={{
                type: DialogType.normal,
                title: strings.Confirmation,
            }}
            modalProps={{
                isBlocking: true,
            }}
        >
            <div>{message}</div>

            {requiresApproval && (
                <div>
                    <div style={{ marginBlock: '16px' }}>
                        <b>{approvalMessage}</b>
                    </div>
                    <div style={{ marginBottom: '16px' }}>
                        <Checkbox label={strings.ConfirmationLabel} onChange={handleApprovalChange} />
                    </div>
                </div>
            )}
            <DialogFooter
                styles={{
                    actionsRight: {
                        display: 'flex',
                        justifyContent: 'space-between',
                        width: '100%',
                    },
                }}
            >
                <PrimaryButton
                    text={strings.Confirm}
                    onClick={onConfirm}
                    disabled={requiresApproval && !isApproved} // This disables the button by default if approval is needed
                />
                <DefaultButton
                    text={strings.Cancel}
                    onClick={onClose}
                />
            </DialogFooter>
        </Dialog>
    );
};

export default MeetingManagementConfirmation;
