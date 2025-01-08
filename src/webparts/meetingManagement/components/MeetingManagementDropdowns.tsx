import * as React from 'react';
import { Dropdown, Toggle } from '@fluentui/react';
import styles from './MeetingManagement.module.scss'; 
import * as strings from 'MeetingManagementWebPartStrings'



interface IMeetingManagementDropdownsProps {
    meetingTypes: any[];
    categories:any[];
    rooms: any[];
    onTypeFilterChange: (selectedType: string) => void;
    onCategoryFilterChange: (selectedCategory: string) => void;
    onRoomFilterChange: (selectedLocation: string) => void;
    onToggleChange: (checked: boolean) => void;
    selectedTypeFilter: string;
    selectedCategoryFilter: string;
    selectedRoomFilter: string;
    showMyRegistrations: boolean;
    filterByMeetingType: string;
    filterByCategory: string;
    filterByRoom: string;
}

const MeetingManagementDropdowns: React.FC<IMeetingManagementDropdownsProps> = ({
    meetingTypes: meetingTypes,
    categories,
    rooms: rooms,
    onTypeFilterChange,
    onCategoryFilterChange,
    onRoomFilterChange: onLocationFilterChange,
    onToggleChange,
    selectedTypeFilter,
    selectedCategoryFilter,
    selectedRoomFilter: selectedLocationFilter,
    showMyRegistrations,
    filterByMeetingType: filterByMeetingType,
    filterByCategory,
    filterByRoom: filterByRoom
}) => {
    

    const meetingTypeOptions = [{ key: "", text: strings.All }, ...meetingTypes.map(meetingType => ({ key: meetingType.key, text: meetingType.text }))];
    const categoryOptions = [{ key: "", text: strings.All }, ...categories.map(category => ({ key: category.key, text: category.text }))];
    const roomOptions = [{ key: "", text: strings.All }, ...rooms.map(room => ({ key: room.key, text: room.text }))];
 
    return (
        <div className={styles.dropdowns}> 
            {!filterByMeetingType && (
                <Dropdown
                    placeholder={strings.ChooseMeeting}
                    label={strings.ChooseMeetingLabel}
                    options={meetingTypeOptions}
                    selectedKey={selectedTypeFilter}
                    styles={{
                        label: { color: 'inherit' },
                        title: {  minWidth: '150px' } 
                    }}
                    onChange={(ev, selectedItem) => onTypeFilterChange(selectedItem?.key as string)}
                />
            )}

            {/* Conditionally render Category dropdown */}
            {!filterByCategory && (
                <Dropdown
                    placeholder={strings.ChooseCategory}
                    label={strings.ChooseCategoryLabel}
                    options={categoryOptions}
                    selectedKey={selectedCategoryFilter}
                    styles={{
                        label: { color: 'inherit' },
                        title: {  minWidth: '150px' } 
                    }}
                    onChange={(ev, selectedItem) => onCategoryFilterChange(selectedItem?.key as string)}
                />
            )}

            {/* Conditionally render Location dropdown */}
            {!filterByRoom && (
                <Dropdown
                    placeholder={strings.ChooseRoom}
                    label={strings.ChooseRoomLabel}
                    options={roomOptions}
                    selectedKey={selectedLocationFilter}
                    styles={{
                        label: { color: 'inherit' },
                        title: {  minWidth: '150px' } 
                    }}
                    onChange={(ev, selectedItem) => onLocationFilterChange(selectedItem?.key as string)}
                />
            )}

            {/* Toggle to show registrations */}
            <Toggle
                label={strings.ShowMyRegistrations}
                onText={strings.Yes}
                offText={strings.No}
                checked={showMyRegistrations}
                styles={{
                    label: { color: 'inherit' },
                    text: { color: 'inherit' },
                }}
                onChange={(ev, checked) => onToggleChange(checked)}
            />
        </div>
    );
};

export default MeetingManagementDropdowns;
