export interface IPageConfig {
  page: "Dashboard" | "Events" | "Documents" | "Certificates";
  isToastify:string
}

export namespace IEventDetails {
  export interface IEventInfo {
    CourseType: string;
    CourseName: string;
    Title: string;
    Date: Date;
    Members: any[];
    MembersID: any[];
    IsDeleted: Boolean;
    Description: string;
    StartTime: null;
    EndTime: null;
    CourseConfigId: null;
    ID: null;
    supportingDocuments: any[];
    FolderName: string;
    voluntary: any[];
    voluntaryID: any[];
    CreatedBy: string;
    CreatedByID: number | null;
    CreatedByEmail: string;
    modifiedon: string;
    isEventCheck: Boolean;
    allparticipants: any[];
  }

  export interface IFiles {
    name: string;
    url: string;
    type: string;
    file: any[];
    Types: string;
    RefID: Number;
  }

  export interface ISiteUsers {
    key: Number;
    imgUrl: String;
    text: String;
    ID: Number;
    secondaryText: String;
    isValid: Boolean;
  }

  export interface IVideo {
    name: String;
    url: String;
    type: String;
    TimeCreated: String;
  }
  export interface IDocuments {
    FolderName: String;
    supportingDocuments: any[];
  }

  //Upcoming Componet

  export interface IAttendanceModal {
    Id: number;
    isOpen: boolean;
    Members: [];
    selectAll: any;
  }

  export interface IMemberDetails {
    imgUrl: String;

    text: String;
    memberID: Number;
    isAttendance: Number;
    eventID: Number;
    EventAttendanceID: Number;
    isVoluntary: Boolean;
  }
}
