export interface IMeeting {
    Id: number;
    Title: string;
    MeetingType: string;
    Category: string;
    StartTime: string;
    EndTime: string;
    RegistrationDeadline: Date;
    Room: string;
    Description: string;
    //UniqueId: any;
    MaxRegistrations: number;
    //OData__ModernAudienceTargetUserFieldId: number[];
    //Background: any;
    Approval: boolean;
  }