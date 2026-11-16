import * as React from "react";
import styles from "./UpcomingEvents.module.scss";
import { Calendar } from "primereact/calendar";
import { Button } from "primereact/button";
import { InputText } from "primereact/inputtext";
import { Dropdown } from "primereact/dropdown";
import { InputTextarea } from "primereact/inputtextarea";
import { Dialog } from "primereact/dialog";
import { useState } from "react";
import SpServices from "../../../../CommonServices/SpServices";
import { Label, NormalPeoplePicker } from "@fluentui/react";
import * as moment from "moment";
import { RadioButton } from "primereact/radiobutton";

import { TriStateCheckbox } from "primereact/tristatecheckbox";

import { Tooltip } from "primereact/tooltip";
import { IEventDetails } from "../../Config/Interface";

import { sp } from "@pnp/sp";



const EventDetails = (props: any) => {
  const formData: any = {
    Title: "",
    CourseType: "",
    CourseName: "",
    Date: new Date(
      `${props.selectedDate.slice(0, 4)}-${props.selectedDate.slice(
        4,
        6
      )}-${props.selectedDate.slice(6, 8)}`
    ),
    Members: [],
    MembersID: [],
    IsDeleted: false,
    Description: "",
    EndTime: new Date(),
    StartTime: new Date(),
    ID: null,
    CourseConfigId: null,
    supportingDocuments: [],
    oldFile: "",
    courseTypeValue: "Exist",
    courseNameValue: "Exist",
    voluntary: [],
    voluntaryID: [],
    CreatedBy: "",
    CreatedByID: null,
    CreatedByEmail: "",
    isEventCheck: false,
    MeetingLink: "",
  };

  const feedBackData: any = {
    comments: "",
    EventId: null,
    CreatedBy: "",
    CreatedByID: null,
    CreatedByEmail: "",
  };
  interface IFiles {
    name: string;
    url: string;
    type: string;
    file?: any;
    Types: string;
    RefID: Number;
    ID: null;
    owner: Boolean;

    eventIdId: any[];
    baseEventId: null;
  }

  const [EventCurrentValue, SetEventCurrentValue] = React.useState({
    ...formData,
  });
  // const [EventDetails, setEventDetails] = React.useState<any>([]);
  const [curTabEvent, setCurTabEvent] = useState<any>([]);
  const [CourseChoice, setCourseChoice] = useState<any>([]);
  const [CourseListChoice, setCourseListChoice] = useState<any>([]);
  const [formMode, setformMode] = useState<any>([]);
  const [Error, setError] = useState("");
  const [popUp, setPopup] = useState(false);
  //const [count, setCount] = useState(0);

  const [membersDetails, setMembersDetails] = useState<any>([]);
  const [feedbackList, setFeedbackList] = useState<any[]>([]);
  const [currentFeedback, setCurrentFeedback] = useState<any>(feedBackData);
  const [FeedbackText, setFeedbackText] = useState("");

  const [feedBackVisible, setFeedBackVisible] = useState({
    visible: false,
    ID: null,
  });
  const [activeTab, setActiveTab] = useState<string>("TodayEvent");

  const [isAttendanceModal, setisAttendanceModal] =
    useState<IEventDetails.IAttendanceModal>({
      isOpen: false,
      Members: [],
      Id: 0,
      selectAll: null,
    });
  const [isDeletePopup, setisDeletePopup] = useState({
    isOpen: false,
    Id: null,
  });

  const [UserfeedBackVisible, setUserFeedBackVisible] = useState({
    visible: false,
    ID: null,
  });

  // const [isChanged, setIsChanged] = useState<boolean>(false);

  const _toDay = moment().startOf("day"); // Today at 00:00
  const _momentDate = moment(props.selectedDate, "YYYYMMDD").startOf("day");

  const logos = {
    pdf: require("../../components/assets/pdf.png"),
    ppt: require("../../components/assets/ppt.png"),
    excel: require("../../components/assets/xls.png"),
    word: require("../../components/assets/doc.png"),
    image: require("../../components/assets/image.png"),
    another: require("../../components/assets/folder.png"),
  };

  const colorClasses = [
    { bar: styles.tealBar, color: styles.tealcolor },
    { bar: styles.yellowBar, color: styles.yellowcolor },
    { bar: styles.greenBar, color: styles.greencolor },
    { bar: styles.purpleBar, color: styles.purplecolor },
  ];
  function formatTimeToAMPM(timeStr: any) {
    const [hour, minute] = timeStr.split(":");
    const date = new Date();
    date.setHours(+hour);
    date.setMinutes(+minute);
    return date.toLocaleTimeString("en-IN", {
      hour: "numeric",
      minute: "numeric",
      hour12: true,
    });
  }

  const ErrFunction = (err: string, funcName: string): void => {
    console.error(funcName, err);
  };
  const getImageUrl = (url: any) => {
    let imageUrl: any = "";

    if (url) {
      let type: string = url.split(".").pop().toLowerCase();
      if (type == "pdf") {
        imageUrl = logos.pdf;
      } else if (type == "ppt" || type == "pptx" || type == "ppsx") {
        imageUrl = logos.ppt;
      } else if (type == "xls" || type == "xlsx" || type == "csv") {
        imageUrl = logos.excel;
      } else if (type == "docx" || type == "doc") {
        imageUrl = logos.word;
      } else if (
        type == "png" ||
        type == "svg" ||
        type == "jpg" ||
        type == "jpeg"
      ) {
        imageUrl = logos.image;
      } else {
        imageUrl = logos.another;
      }
    } else {
      imageUrl = logos.another;
    }
    return imageUrl;
  };

  const today = new Date();
  const todayFormatted = `${today.getFullYear()}${String(
    today.getMonth() + 1
  ).padStart(2, "0")}`;

  // function isPast(dateString: string, timeString: string): boolean {
  //   console.log(dateString, "date");
  //   const eventDateTime = new Date(
  //     `${dateString.split("T")[0]}T${timeString}:00`
  //   );
  //   const currentDateTime = new Date();

  //   return eventDateTime.getTime() < currentDateTime.getTime();
  // }

  function isPast(dateValue: Date, timeString: string): boolean {
    const eventDateTime = new Date(dateValue);

    if (timeString) {
      const [hours, minutes] = timeString.split(":").map(Number);
      eventDateTime.setHours(hours, minutes, 0, 0);
    }

    return eventDateTime.getTime() < Date.now();
  }

  const checkDuplicateTitleIgnoreCase = async (
    title: string
  ): Promise<boolean> => {
    try {
      const cleanInput = title;

      const items = await sp.web.lists
        .getByTitle("Event")
        .items.select("Title,IsDeleted")
        .top(5000)();

      const match = items.find(
        (item) =>
          item.IsDeleted === false &&
          item.Title?.replace(/\s+/g, "").toLowerCase() ===
            cleanInput.replace(/\s+/g, "").toLowerCase()
      );

      return !!match;
    } catch (err) {
      console.error("Error checking title:", err);
      return false;
    }
  };

  const parseTimeStringToDate = (timeString: string): Date => {
    const [hours, minutes] = timeString.split(":").map(Number);
    const date = new Date();
    date.setHours(hours, minutes, 0, 0);
    return date;
  };
  function _sortFilterKeys(a: any, b: any) {
    if (a.modifiedon > b.modifiedon) {
      return -1;
    }
    if (a.modifiedon < b.modifiedon) {
      return 1;
    }
    return 0;
  }

  // const fetchEvents = async () => {
  //   const startDate = moment(props.selectedDate, "YYYYMMDD")
  //     .startOf("day")
  //     .toISOString();
  //   const endDate = moment(props.selectedDate, "YYYYMMDD")
  //     .endOf("day")
  //     .toISOString();

  //   await SpServices.SPReadItems({
  //     Listname: "Event",
  //     Select:
  //       "*,Members/Title,Members/ID,Members/EMail,voluntary/Title,voluntary/ID,voluntary/EMail,Author/Title,Author/ID,Author/EMail",
  //     Expand: "Members,voluntary,Author",
  //     Orderby: "Modified",
  //     Filter: [
  //       {
  //         FilterKey: "Date",
  //         Operator: "ge",
  //         FilterValue: startDate,
  //       },
  //       {
  //         FilterKey: "Date",
  //         Operator: "le",
  //         FilterValue: endDate,
  //       },
  //       {
  //         FilterKey: "IsDeleted",
  //         Operator: "eq",
  //         FilterValue: "0",
  //       },
  //     ],
  //   })
  //     .then((events: any) => {
  //       let _EventData: any = [];
  //       console.log(events, "events");
  //       if (events.length) {
  //         let _count = 0;
  //         for (let i = 0; i < events.length; i++) {
  //           _count++;
  //           const event = events[i];
  //           void SpServices.SPReadItemUsingId({
  //             Listname: "CourseConfig",
  //             SelectedId: event?.CourseConfigId || null,
  //             Select: "*",
  //             Expand: "",
  //           })
  //             .then(async (CourseItem: any) => {
  //               _EventData.push({
  //                 CourseType: CourseItem.CourseType || "",
  //                 CourseName: CourseItem.CourseName || "",
  //                 Title: event.Title || "",
  //                 Date: event.Date || new Date(),
  //                 Members: event.Members?.length ? event.Members : [],
  //                 MembersID: event.MembersId?.length ? event.MembersId : [],
  //                 IsDeleted: false,
  //                 Description: event.Description || "",
  //                 EndTime: event.EndTime || "",
  //                 StartTime: event.StartTime || "",
  //                 CourseConfigId: event.CourseConfigId || null,
  //                 ID: event.ID || null,
  //                 voluntary: event.voluntary?.length ? event.voluntary : [],
  //                 voluntaryID: event.voluntaryId?.length
  //                   ? event.voluntaryId
  //                   : [],
  //                 CreatedBy: event.Author?.Title || "",
  //                 CreatedByID: event.Author?.ID || null,
  //                 CreatedByEmail: event.Author?.EMail || "",
  //                 modifiedon: event.Modified,
  //                 allparticipants: [
  //                   ...(event.voluntary?.length ? event.voluntary : []),
  //                   ...(event.Members?.length ? event.Members : []),
  //                 ],
  //                 isEventCheck: false,
  //               });

  //               if (events.length === _count) {
  //                 setEventDetails([..._EventData]);
  //                 console.log(_EventData, "EventDetails");
  //                 filterAndSetEventData([..._EventData], "TodayEvent");
  //                 // setLoaderValue(false);
  //               }
  //             })
  //             .catch((err: any) => {
  //               ErrFunction("Event Attachment Data get error", err);
  //             });
  //         }
  //       } else {
  //         setEventDetails([]);
  //         filterAndSetEventData([], "TodayEvent");
  //       }
  //     })
  //     .catch((err) => {
  //       ErrFunction(err, "Get CourseConfig");
  //     })
  //     .catch((err: any) => {
  //       console.error("Error fetching events", err);
  //     });
  // };

  function filterAndSetEventData(EventDatas: any[], eventType: string) {
    let filteredEvents: any[] = [];
    const now = moment();
    const selectedDate = props.selectedDate
      ? moment(props.selectedDate).format("YYYYMMDD")
      : null;

    if (eventType === "TodayEvent" && selectedDate) {
      filteredEvents = EventDatas.filter((event: any) => {
        const eventDate = moment(event.Date).format("YYYYMMDD");
        return eventDate === selectedDate;
      }).sort(_sortFilterKeys);
    } else {
      filteredEvents = EventDatas.filter((event: any) => {
        const eventDate = moment(event.Date);
        return eventDate.isAfter(now, "day");
      })
        .sort(_sortFilterKeys)
        .slice(0, 10);
    }

    setCurTabEvent(filteredEvents);

    setActiveTab(eventType);
    void GetAllChoices();
  }

  // function filterAndSetEventData(
  //   EventDatas: any[],
  //   eventType: string,
  //   selectedDate?: Date
  // ) {
  //   let filteredEvents: any[] = [];
  //   const now = new Date();

  //   const normalize = (date: Date) =>
  //     new Date(date.getFullYear(), date.getMonth(), date.getDate());

  //   if (selectedDate) {
  //     const selDate = normalize(selectedDate);

  //     filteredEvents = EventDatas.filter((event: any) => {
  //       const d = normalize(new Date(event.Date));
  //       return d.getTime() === selDate.getTime();
  //     }).sort(_sortFilterKeys);
  //   } else if (eventType === "MyEvents") {
  //     filteredEvents = EventDatas.filter(
  //       (event: any) => new Date(event.Date) >= now
  //     )
  //       .sort(_sortFilterKeys)
  //       .slice(0, 10);
  //   } else {
  //     const today = normalize(now);

  //     filteredEvents = EventDatas.filter((event: any) => {
  //       const d = normalize(new Date(event.Date));
  //       return d.getTime() === today.getTime();
  //     }).sort(_sortFilterKeys);
  //   }

  //   setCurTabEvent(filteredEvents);
  //   setActiveTab(eventType);
  //   void GetAllChoices();

  //   console.log(`${eventType} →`, filteredEvents);
  // }

  function GetAllChoices() {
    sp.web.lists
      .getByTitle("CourseConfig")
      .fields.get()
      .then((choices) => {
        let _typeChoiceStatus: any[] = choices.filter(
          (e) => e.StaticName == "CourseType"
        );
        let CourseType: any = [];
        if (_typeChoiceStatus.length) {
          _typeChoiceStatus[0].Choices.forEach((item: string) => {
            if (!CourseType.some((choice: any) => choice.name == item)) {
              CourseType.push({
                name: item || "",
                code: item || "",
              });
            }
          });
        }
        setCourseChoice(CourseType);
        void GetCourseConfig();
      })
      .catch((err: any) => {
        ErrFunction("Get Choices error", err);
      });
  }
  function GetCourseConfig() {
    SpServices.SPReadItems({
      Listname: "CourseConfig",
    })
      .then((res: any) => {
        let _CourseType: any = [];
        if (res.length) {
          res.forEach((item: any) => {
            _CourseType.push({
              name: item.CourseName,
              code: item.CourseName,
              ID: item.ID,
              CourseType: item.CourseType,
            });
          });
        }
        setCourseListChoice([..._CourseType]);
      })
      .catch((err) => ErrFunction(err, "Get Currency"));
  }

  const GetEventFeedBack = (data: any) => {
    sp.web.lists
      .getByTitle("Feedback")
      .items.select("*", "Author/Title", "Author/ID", "Author/EMail")
      .expand("Author")
      .filter(`EventId eq '${data.ID}'`)
      .get()
      .then((items: any) => {
        const commentsDetails = items.map((item: any) => ({
          comments: item.Comments,
          EventId: item.EventIdId,
          CreatedBy: item.Author?.Title || "",
          CreatedByID: item.Author?.ID || null,
          CreatedByEmail: item.Author?.EMail || "",
          CreatedDate: item.Created,
        }));

        const currentUserEmail = props.context.pageContext.user.email;

        const matchedFeedbacks = commentsDetails.filter(
          (comment: any) =>
            comment.CreatedByEmail == currentUserEmail &&
            comment?.EventId === data?.ID
        );

        // Get the first match if it exists
        const existingFeedback =
          matchedFeedbacks.length > 0 ? matchedFeedbacks[0] : null;

        // Now update state
        setFeedbackList(commentsDetails);
        setCurrentFeedback(existingFeedback);
        setFeedbackText(existingFeedback ? existingFeedback.comments : "");
      })
      .catch((err: any) => {
        ErrFunction("Get Feedback error", err);
      });
  };

  const getFile = (fileList: any) => {
    let _files: IFiles[] = [...EventCurrentValue.supportingDocuments];

    for (let i = 0; i < fileList.length; i++) {
      const newFile = fileList[i];
      const fileName = newFile.name;

      const isAlreadyExists = _files.filter(
        (f) => f.name === fileName && f.Types !== "IsDelete"
      );

      if (isAlreadyExists.length > 0) {
        setError("The file already exist");
      } else {
        _files.push({
          name: newFile.name,
          url: URL.createObjectURL(newFile),
          type: newFile.name.split(".").pop() || "",
          Types: "Insert",
          RefID: _files.length,
          file: newFile,
          ID: null,
          owner: true,
          eventIdId: EventCurrentValue.ID != null ? [EventCurrentValue.ID] : [],
          baseEventId: null,
        });
      }
    }

    let Filecount = _files.filter((_a) => _a.Types !== "IsDelete").length;

    let fileCountElem = document.getElementById("fileCount");
    if (fileCountElem) {
      fileCountElem.textContent =
        Filecount > 0 ? `${Filecount} file(s)` : "No files";
    }

    SetEventCurrentValue({
      ...EventCurrentValue,
      supportingDocuments: _files,
    });
  };

  const fileDelete = (index: number) => {
    const temp = [...EventCurrentValue.supportingDocuments];
    const currentFile = temp[index];
    if (!currentFile) return;

    if (currentFile.Types === "Inlist") {
      if (currentFile.owner && formMode !== "New") {
        if (currentFile.baseEventId === EventCurrentValue.ID) {
          temp[index].Types = "IsDelete";
        } else {
          temp[index].eventIdId = currentFile.eventIdId?.filter(
            (id: number) => id !== EventCurrentValue.ID
          );
        }
      } else {
        temp.splice(index, 1);
      }
    } else if (currentFile.Types === "Insert") {
      temp.splice(index, 1);
    }
    const Filecount = temp.filter(
      (_a: any) =>
        _a.Types !== "IsDelete" &&
        Array.isArray(_a.eventIdId) &&
        _a.eventIdId.includes(EventCurrentValue.ID)
    ).length;

    const fileCountElem = document.getElementById("fileCount");
    if (fileCountElem) {
      fileCountElem.textContent =
        Filecount > 0 ? `${Filecount} file(s)` : "No files";
    }

    SetEventCurrentValue({
      ...EventCurrentValue,
      supportingDocuments: temp,
    });
  };

  function sanitizeFileName(name: string): string {
    if (!name) return ""; // handle null/undefined
    return name.replace(/[~#%&*{}\\<>?/+|"]/g, "");
  }

  const AddFileDocument = (
    DeleteDocuments: any[],
    InsertDocuments: any[],
    ID: number,
    Path: string
  ) => {
    if (InsertDocuments.length) {
      let insertCount = 0;

      InsertDocuments.forEach((doc) => {
        const fileName = sanitizeFileName(doc.name);
        const fileContent = doc.content;

        sp.web
          .getFolderByServerRelativeUrl(Path)
          .files.add(fileName, fileContent, true)
          .then((addedFile: any) => addedFile.file.getItem())
          .then((item: any) => item.get())
          .then((details: any) => {
            return sp.web.lists
              .getByTitle("Documents")
              .items.getById(details.Id)
              .update({
                EventIdId: { results: [ID] },
                BaseEventId: ID || null,
              });
          })
          .then(() => {
            insertCount++;

            if (insertCount === InsertDocuments.length) {
              const inlistFiles = EventCurrentValue.supportingDocuments.filter(
                (e: any) => e.Types === "Inlist"
              );

              let inlistCount = 0;

              if (inlistFiles.length) {
                inlistFiles.forEach((file: any) => {
                  const eventIds =
                    formMode === "New"
                      ? [...(file.eventIdId || []), ID]
                      : file.eventIdId || [];

                  sp.web.lists
                    .getByTitle("Documents")
                    .items.getById(file.ID)
                    .update({ EventIdId: { results: eventIds } })
                    .then(() => {
                      inlistCount++;
                      if (inlistCount === inlistFiles.length) {
                        setPopup(false);
                        props.setPageConfig({
                          page: "Dashboard",
                          isToastify: formMode === "New" ? "Success" : "Info",
                        });
                      }
                    })
                    .catch((err) => ErrFunction("Update Inlist error", err));
                });
              } else {
                setPopup(false);
                props.setPageConfig({
                  page: "Dashboard",
                  isToastify: formMode === "New" ? "Success" : "Info",
                });
              }
            }
          })
          .catch((err) => ErrFunction("AddFileDocument error", err));
      });
    } else {
      const inlistFiles = EventCurrentValue.supportingDocuments.filter(
        (e: any) => e.Types === "Inlist"
      );

      let inlistCount = 0;

      if (inlistFiles.length) {
        inlistFiles.forEach((file: any) => {
          const eventIds =
            formMode === "New"
              ? [...(file.eventIdId || []), ID]
              : file.eventIdId || [];

          sp.web.lists
            .getByTitle("Documents")
            .items.getById(file.ID)
            .update({ EventIdId: { results: eventIds } })
            .then(() => {
              inlistCount++;
              if (inlistCount === inlistFiles.length) {
                setPopup(false);
                props.setPageConfig({
                  page: "Dashboard",
                  isToastify: formMode === "New" ? "Success" : "Info",
                });
              }
            })
            .catch((err) => ErrFunction("Update Inlist error", err));
        });
      } else {
        setPopup(false);
        props.setPageConfig({
          page: "Dashboard",
          isToastify: formMode === "New" ? "Success" : "Info",
        });
      }
    }
  };

  const DeleteFileDocument = (
    DeleteDocuments: any[],
    InsertDocuments: any[],
    ID: number,
    Path: string
  ) => {
    if (DeleteDocuments.length > 0) {
      let deleteCount = 0;

      DeleteDocuments.forEach((delDoc) => {
        const fileName = sanitizeFileName(delDoc.name);

        sp.web
          .getFolderByServerRelativeUrl(Path)
          .files.getByName(fileName)
          .delete()
          .then(() => {
            deleteCount++;
            if (deleteCount === DeleteDocuments.length) {
              AddFileDocument(DeleteDocuments, InsertDocuments, ID, Path);
            }
          })
          .catch((err) => ErrFunction("Document delete error", err));
      });
    } else {
      AddFileDocument(DeleteDocuments, InsertDocuments, ID, Path);
    }
  };

  const AddFolderDocument = (
    DeleteDocuments: any[],
    InsertDocuments: any[],
    ID: number,
    Path: ""
  ) => {
    try {
      const list = sp.web.lists.getByTitle("Documents");

      const courseTypeName =
        EventCurrentValue.CourseType?.trim() || "DefaultCourseType";
      const courseName =
        EventCurrentValue.CourseName?.trim() || "DefaultCourseName";
      const courseTitle = EventCurrentValue.Title?.trim() || "DefaultTitle";

      // Step 1: Get root folders (Shared Documents)
      list.rootFolder.folders
        .select("*")
        .get()
        .then((rootFolders: any[]) => {
          let courseTypeFolder = rootFolders.find(
            (f) => f.Name === courseTypeName
          );

          const createCourseNameFolder = (courseTypeUrl: string) => {
            sp.web
              .getFolderByServerRelativeUrl(courseTypeUrl)
              .folders.select("*")
              .get()
              .then((subFolders: any[]) => {
                let courseNameFolder = subFolders.find(
                  (f) => f.Name === courseName
                );

                const createCourseTitleFolder = (courseNameUrl: string) => {
                  sp.web
                    .getFolderByServerRelativeUrl(courseNameUrl)
                    .folders.select("*")
                    .get()
                    .then((titleFolders: any[]) => {
                      let courseTitleFolder = titleFolders.find(
                        (f) => f.Name === courseTitle
                      );

                      if (courseTitleFolder) {
                        DeleteFileDocument(
                          DeleteDocuments,
                          InsertDocuments,
                          ID,
                          courseTitleFolder.ServerRelativeUrl
                        );
                      } else {
                        sp.web
                          .getFolderByServerRelativeUrl(courseNameUrl)
                          .folders.add(courseTitle)
                          .then((newTitle) => {
                            DeleteFileDocument(
                              DeleteDocuments,
                              InsertDocuments,
                              ID,
                              newTitle.data.ServerRelativeUrl
                            );
                          })
                          .catch((err) =>
                            ErrFunction(
                              "CourseTitle folder creation error",
                              err
                            )
                          );
                      }
                    })
                    .catch((err) =>
                      ErrFunction("CourseTitle subfolder check error", err)
                    );
                };

                if (courseNameFolder) {
                  createCourseTitleFolder(courseNameFolder.ServerRelativeUrl);
                } else {
                  sp.web
                    .getFolderByServerRelativeUrl(courseTypeUrl)
                    .folders.add(courseName)
                    .then((newFolder) => {
                      createCourseTitleFolder(newFolder.data.ServerRelativeUrl);
                    })
                    .catch((err) =>
                      ErrFunction("CourseName folder creation error", err)
                    );
                }
              })
              .catch((err) =>
                ErrFunction("CourseName subfolder check error", err)
              );
          };

          // Step 2: Create CourseType folder if needed
          if (courseTypeFolder) {
            createCourseNameFolder(courseTypeFolder.ServerRelativeUrl);
          } else {
            list.rootFolder.folders
              .add(courseTypeName)
              .then((newCourseType) => {
                createCourseNameFolder(newCourseType.data.ServerRelativeUrl);
              })
              .catch((err) =>
                ErrFunction("CourseType folder creation error", err)
              );
          }
        })
        .catch((err) => ErrFunction("Root folder check error", err));
    } catch (err) {
      ErrFunction("AddFolderDocument general error", err);
    }
  };

  const fileAddandUpdate = (ID: number, formMode: string) => {
    try {
      const supportingDocuments = EventCurrentValue.supportingDocuments;
      const InsertDocuments: any[] = [];
      const DeleteDocuments: any[] = [];

      supportingDocuments.forEach((doc: any) => {
        if (doc.Types === "Insert") {
          InsertDocuments.push({ name: doc.name, content: doc.file });
        } else if (doc.Types === "IsDelete") {
          DeleteDocuments.push({ name: doc.name, content: doc.file });
        }
      });

      const list = sp.web.lists.getByTitle("Documents");

      const folderPath = `${props.SiteName}/Shared Documents/${EventCurrentValue.CourseType}/${EventCurrentValue.CourseName}/${EventCurrentValue.Title}`;

      list.rootFolder.folders
        .select("*")
        .get()
        .then((rootFolders: any[]) => {
          const courseTypeFolder = rootFolders.find(
            (f) => f.Name === EventCurrentValue.CourseType
          );

          if (courseTypeFolder) {
            sp.web
              .getFolderByServerRelativeUrl(courseTypeFolder.ServerRelativeUrl)
              .folders.select("*")
              .get()
              .then((courseNameFolders: any[]) => {
                const courseNameFolder = courseNameFolders.find(
                  (f) => f.Name === EventCurrentValue.CourseName
                );

                if (courseNameFolder) {
                  sp.web
                    .getFolderByServerRelativeUrl(
                      courseNameFolder.ServerRelativeUrl
                    )
                    .folders.select("*")
                    .get()
                    .then((courseTitleFolders: any[]) => {
                      const courseTitleFolder = courseTitleFolders.find(
                        (f) => f.Name === EventCurrentValue.Title
                      );

                      if (courseTitleFolder) {
                        DeleteFileDocument(
                          DeleteDocuments,
                          InsertDocuments,
                          ID,
                          folderPath
                        );
                      } else {
                        AddFolderDocument(
                          DeleteDocuments,
                          InsertDocuments,
                          ID,
                          ""
                        );
                      }
                    })
                    .catch((err) =>
                      ErrFunction("CourseTitle subfolder fetch error", err)
                    );
                } else {
                  AddFolderDocument(DeleteDocuments, InsertDocuments, ID, "");
                }
              })
              .catch((err) =>
                ErrFunction("CourseName subfolder fetch error", err)
              );
          } else {
            AddFolderDocument(DeleteDocuments, InsertDocuments, ID, "");
          }
        })
        .catch((err) => ErrFunction("Root folder fetch error", err));
    } catch (err) {
      ErrFunction("fileAddandUpdate general error", err);
    }
  };

  const getAllFilesRecursive = async (folderUrl: string): Promise<any[]> => {
    let allFiles: any[] = [];

    try {
      // 1️⃣ Get current folder's files
      const files = await sp.web
        .getFolderByServerRelativeUrl(folderUrl)
        .files.select(
          "Name",
          "ServerRelativeUrl",
          "TimeLastModified",
          "ListItemAllFields/Id",
          "ListItemAllFields/EventIdId",
          "ListItemAllFields/BaseEventId"
        )
        .expand("ListItemAllFields")
        .get();

      allFiles.push(...files);

      const subFolders = await sp.web
        .getFolderByServerRelativeUrl(folderUrl)
        .folders.select("Name", "ServerRelativeUrl")
        .get();

      for (const sub of subFolders) {
        const subFiles = await getAllFilesRecursive(sub.ServerRelativeUrl);
        allFiles.push(...subFiles);
      }
    } catch (err) {
      console.error("Error fetching files recursively:", err);
    }

    return allFiles;
  };

  const getEvent = (EventID: any, formMode: string) => {
    void SpServices.SPReadItemUsingId({
      Listname: "Event",
      SelectedId: EventID,
      Select: "*,Members/Title,Members/ID,Members/EMail",
      Expand: "Members",
    }).then(async (item: any) => {
      let _attachments: IFiles[] = [];
      let _VideoUrl: string | null = null;

      await SpServices.SPReadItemUsingId({
        Listname: "CourseConfig",
        SelectedId: item?.CourseConfigId || null,
        Select: "*",
        Expand: "",
      }).then(async (CourseItem: any) => {
        const folderUrl = `${props.SiteName}/Shared Documents/${CourseItem.CourseType}/${CourseItem.CourseName}`;

        const allFiles = await getAllFilesRecursive(folderUrl);
        const attachmentData = allFiles.filter((f: any) =>
          f.ListItemAllFields?.EventIdId
            ? (f.ListItemAllFields.EventIdId as number[]).includes(item.ID)
            : false
        );
        if (attachmentData.length) {
          attachmentData.forEach((file: any) => {
            _attachments.push({
              name: file.Name,
              url: file.ServerRelativeUrl,
              type: file.Name.split(".").pop(),
              Types: "Inlist",
              RefID: item.ID + 1,
              ID: file.ListItemAllFields.ID,
              owner: true,
              baseEventId: file.ListItemAllFields.BaseEventId || null,
              eventIdId: file.ListItemAllFields.EventIdId || [],
            });
          });
        }

        SpServices.SPReadItems({
          Listname: "RecordingVideos",
          Filter: [
            {
              FilterKey: "EventId",
              Operator: "eq",
              FilterValue: item.ID,
            },
          ],
          Select: "*,File/Name,File/ServerRelativeUrl",
          Expand: "File",
        })

          .then((VideoData: any[]) => {
            if (VideoData.length) {
              const video = VideoData.find((v) => v.File);
              if (video) _VideoUrl = video.File.ServerRelativeUrl;
            }

            // Prepare final Event object
            const _d: any = {
              CourseType: CourseItem.CourseType || "",
              CourseName: CourseItem.CourseName || "",
              Title: item.Title || "",
              Date: item.Date || new Date(),
              Members: item.Members?.length ? item.Members : [],
              MembersID: item.MembersId?.length ? item.MembersId : [],
              IsDeleted: item.IsDeleted || false,
              Description: item.Description || "",
              EndTime: parseTimeStringToDate(item.EndTime) || null,
              StartTime: parseTimeStringToDate(item.StartTime) || null,
              CourseConfigId: item.CourseConfigId || null,
              supportingDocuments: _attachments,
              ID: item.ID || null,
              oldFile: item.Title,
              courseTypeValue: "Exist",
              courseNameValue: "Exist",
              voluntary: item.voluntary?.length ? item.voluntary : [],
              voluntaryID: item.voluntaryId?.length ? item.voluntaryId : [],
              CreatedBy: item.Author?.Title || "",
              CreatedByID: item.Author?.ID || null,
              CreatedByEmail: item.Author?.EMail || "",
              VideoUrl: _VideoUrl || null,
              isEventCheck: false,
              MeetingLink: item.MeetingLink || "",
            };

            let Filecount = _attachments.filter(
              (_a) => _a.Types !== "IsDelete"
            ).length;

            let fileCountElem = document.getElementById("fileCount");
            if (fileCountElem) {
              fileCountElem.textContent =
                Filecount > 0 ? `${Filecount} file(s)` : "No files";
            }

            SetEventCurrentValue(_d);
            setformMode(formMode);
            setPopup(true);
          })
          .catch((err: any) => {
            ErrFunction("Get Recording Video error", err);
          })
          .catch((err: any) => {
            ErrFunction("Get CourseConfig error", err);
          })
          .catch((err: any) => {
            ErrFunction("Get CourseConfig error", err);
          });
      });
    });
  };

  const findUserAll = (user: any) => {
    let finduser: any = [];
    if (user.length) {
      user.forEach((userId: any) => {
        let user = [...props.SiteUsers].filter((val: any) => {
          return val.ID === userId;
        });
        finduser.push(user[0]);
      });
    }
    return finduser;
  };

  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };

  const GetUserDetails = (filterText: any, currentPersonas: any) => {
    let _allPeoples = [...props.SiteUsers];

    _allPeoples = _allPeoples.filter((arr) => {
      return arr.text.toLowerCase().indexOf("archive") == -1;
    });

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (arr) => !currentPersonas.some((persona: any) => persona.ID == arr.ID)
      );
    }

    let result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };

  function getPreviousFiles(CourseName: string): Promise<any[]> {
    const _attachments: any[] = [];

    return sp.web
      .getFolderByServerRelativeUrl(`${props.SiteName}/Shared Documents`)
      .folders()
      .then((rootFolders: any[]) => {
        const promises = rootFolders.map((typeFolder) => {
          return sp.web
            .getFolderByServerRelativeUrl(typeFolder.ServerRelativeUrl)
            .folders()
            .then((courseNameFolders: any[]) => {
              const courseFolder = courseNameFolders.find(
                (f: any) =>
                  f.Name?.toLowerCase().trim() ===
                  CourseName.toLowerCase().trim()
              );

              if (!courseFolder) return Promise.resolve(); // Skip if not found

              return sp.web
                .getFolderByServerRelativeUrl(courseFolder.ServerRelativeUrl)
                .folders()
                .then((courseTitleFolders: any[]) => {
                  const titlePromises = courseTitleFolders.map(
                    (titleFolder) => {
                      return sp.web
                        .getFolderByServerRelativeUrl(
                          titleFolder.ServerRelativeUrl
                        )
                        .files.select(
                          "FileLeafRef,Name,ServerRelativeUrl,ListItemAllFields/Id,ListItemAllFields/EventIdId,ListItemAllFields/BaseEventId"
                        )
                        .expand("ListItemAllFields")
                        .get()
                        .then((_attachmentData: any[]) => {
                          const files = EventCurrentValue?.ID
                            ? _attachmentData.filter((f) =>
                                f.ListItemAllFields.EventIdId?.includes(
                                  EventCurrentValue.ID
                                )
                              )
                            : _attachmentData;

                          files.forEach((file) => {
                            _attachments.push({
                              name: file.Name,
                              url: file.ServerRelativeUrl,
                              type: file.Name.split(".").pop() || "",
                              Types: "Inlist",
                              RefID: 1,
                              ID: file.ListItemAllFields?.ID || null,
                              baseEventId:
                                file.ListItemAllFields?.BaseEventId || null,
                              eventIdId:
                                file.ListItemAllFields?.EventIdId || [],
                              owner: EventCurrentValue?.ID
                                ? file.ListItemAllFields?.EventIdId?.includes(
                                    EventCurrentValue.ID
                                  )
                                : false,
                            });
                          });
                        })
                        .catch((err) =>
                          ErrFunction(
                            "Error fetching files from title folder",
                            err
                          )
                        );
                    }
                  );

                  return Promise.all(titlePromises);
                })
                .catch((err) =>
                  ErrFunction("Error fetching courseTitle folders", err)
                );
            })
            .catch((err) =>
              ErrFunction("Error fetching courseName folders", err)
            );
        });

        return Promise.all(promises);
      })
      .then(() => {
        console.log("✅ Total attachments found:", _attachments.length);
        return _attachments;
      })
      .catch((err) => {
        console.error("❌ Error fetching previous files:", err);
        ErrFunction("Event Attachment Data get error", err.message || err);
        return [];
      });
  }

  const handleOnchange = async (key: any, value: any) => {
    let _data: any = { ...EventCurrentValue };

    if (key === "Title") {
      _data["Title"] = value;
    }

    if (key === "MembersID") {
      let Membersid: any = [];
      value.forEach((val: any) => Membersid.push(val.ID));
      _data[key] = Membersid;
    } else if (key === "CourseName") {
      if (_data.courseNameValue === "New") {
        _data["CourseName"] = value;
        _data["CourseConfigId"] = null;
      } else {
        _data["CourseName"] = value.name;
        _data["CourseConfigId"] = value.ID;
        _data["supportingDocuments"] = [];

        const attachments = await getPreviousFiles(value.name);

        console.log("Attachments from previous course:", attachments);
        _data["supportingDocuments"] = attachments;
      }
    } else if (key === "CourseType") {
      if (_data.courseTypeValue === "New") {
        _data["CourseType"] = value;
      } else {
        _data["CourseType"] = value.name;
      }
    } else if (key == "courseTypeValue") {
      _data["CourseType"] = "";
      _data["CourseName"] = "";
      _data["CourseConfigId"] = null;
      _data["supportingDocuments"] = [];
      _data[key] = value;
    } else if (key == "courseNameValue") {
      _data["CourseName"] = "";
      _data["supportingDocuments"] = [];
      _data[key] = value;
    } else if (key === "Date") {
      props._sDate = moment(value).format("YYYYMMDD");
      let EventDate: any = new Date(value);
      EventDate.setHours(
        value.getHours(),
        value.getMinutes(),
        value.getSeconds()
      );

      _data[key] = value;
      _data["StartTime"] = EventDate;
      _data["EndTime"] = new Date(EventDate.getTime() + 30 * 60000);
    } else if (key === "StartTime") {
      let EventDate: any = new Date(EventCurrentValue.Date);
      EventDate.setHours(
        value.getHours(),
        value.getMinutes(),
        value.getSeconds()
      );

      _data[key] = EventDate;
      _data["EndTime"] = new Date(EventDate.getTime() + 30 * 60000);
    } else if (key === "EndTime") {
      let EventDate: any = new Date(EventCurrentValue.Date);
      EventDate.setHours(
        value.getHours(),
        value.getMinutes(),
        value.getSeconds()
      );
      _data["EndTime"] = value;
    } else {
      _data[key] = value;
    }

    SetEventCurrentValue(_data);
  };

  const handleSubmit = async (ConfigID: any) => {
    let EventJson = {
      Title: EventCurrentValue.Title,
      Date: EventCurrentValue.Date,
      MembersId: {
        results: EventCurrentValue.MembersID ? EventCurrentValue.MembersID : [],
      },
      IsDeleted: EventCurrentValue.IsDeleted,
      Description: EventCurrentValue.Description,
      EndTime: new Date(EventCurrentValue.EndTime).toTimeString().slice(0, 5),
      StartTime: new Date(EventCurrentValue.StartTime)
        .toTimeString()
        .slice(0, 5),

      CourseConfigId:
        ConfigID === null ? EventCurrentValue.CourseConfigId : ConfigID,
      isEventCheck: true,
    };

    if (formMode === "New") {
      props.setLoaderValue(true);
      SpServices.SPAddItem({
        Listname: "Event",
        RequestJSON: EventJson,
      })
        .then((res) => {
          if (EventCurrentValue.supportingDocuments.length > 0) {
            void fileAddandUpdate(res.data.ID, "New");
          } else {
            setPopup(false);

            props.setPageConfig({ page: "Dashboard", isToastify: "Success" });
          }
        })
        .catch((err: any) => {
          ErrFunction("Employee Master add data error", err);
        });
    } else {
      props.setLoaderValue(true);

      SpServices.SPUpdateItem({
        Listname: "Event",
        ID: Number(EventCurrentValue.ID),
        RequestJSON: EventJson,
      })
        .then((res) => {
          if (
            EventCurrentValue.supportingDocuments.filter(
              (e: any) => e.Type !== "Inlist"
            ).length > 0
          ) {
            void fileAddandUpdate(EventCurrentValue.ID, "Edit");
          } else {
            setPopup(false);
            props.setPageConfig({ page: "Dashboard", isToastify: "Info" });
          }
        })
        .catch((err: any) => {
          ErrFunction("Employee Master add data error", err);
        });
    }
  };

  const newCourseDetails = async () => {
    let CourseJson = {
      CourseType: EventCurrentValue.CourseType || "",
      CourseName: EventCurrentValue.CourseName || "",
    };

    try {
      const field: any = await sp.web.lists
        .getByTitle("CourseConfig")
        .fields.getByInternalNameOrTitle("CourseType")
        .get();

      const choices = [...field.Choices];

      if (!choices.includes(EventCurrentValue.CourseType)) {
        choices.push(EventCurrentValue.CourseType);

        await sp.web.lists
          .getByTitle("CourseConfig")
          .fields.getByInternalNameOrTitle("CourseType")
          .update({
            Choices: { results: choices }, // Proper array structure
          });
      }

      // Add new item to the list
      const res = await SpServices.SPAddItem({
        Listname: "CourseConfig",
        RequestJSON: CourseJson,
      });

      void handleSubmit(res.data.ID);
    } catch (err) {
      ErrFunction("new courseDetails error", err);
    }
  };

  const Validation = async (event: any) => {
    let error = "";

    // Title validation

    if (!event.Title || event.Title.trim() === "") {
      error = "Please enter a title";
    } else if (
      formMode === "New" &&
      (await checkDuplicateTitleIgnoreCase(event.Title || ""))
    ) {
      error =
        "An event with this title already exists. Please enter another title";
    }

    // Course type validation
    else if (!event.CourseType?.trim() && event.courseTypeValue === "New") {
      error = "Please enter a course type";
    } else if (!event.CourseType?.trim()) {
      error = "Please choose the type of course";
    } else if (
      event.courseTypeValue === "New" &&
      CourseChoice.some(
        (e: any) =>
          e.name?.replace(/\s+/g, "").toLowerCase() ===
          event.CourseType.replace(/\s+/g, "").toLowerCase()
      )
    ) {
      error = "A course with this type already exists";
    }

    // Course name validation
    else if (!event.CourseName?.trim() && event.courseNameValue === "New") {
      error = "Please enter a course name";
    } else if (!event.CourseName?.trim()) {
      error = "Please choose the name of course";
    } else if (
      event.courseNameValue === "New" &&
      CourseListChoice.some(
        (e: any) =>
          e.name?.replace(/\s+/g, "").toLowerCase() ===
          event.CourseName.replace(/\s+/g, "").toLowerCase()
      )
    ) {
      error = "This course name is already in use";
    }

    // Date validation
    else if (!event.Date) {
      error = "Please pick a date";
    }

    // Time validation
    else if (!event.StartTime) {
      error = "Please pick a start time";
    } else if (!event.EndTime) {
      error = "Please pick a end time";
    } else if (
      moment(event.Date).isSame(moment(), "day") &&
      moment(event.StartTime).isBefore(moment(), "minute") &&
      formMode === "New"
    ) {
      error = "This start time has already passed. Pick a future time";
    } else if (
      new Date(event.StartTime).getTime() === new Date(event.EndTime).getTime()
    ) {
      error = "Start and end times must be different";
    }

    // Members validation
    else if (!event.MembersID || event.MembersID.length === 0) {
      error = "At least one member is required";
    }

    // No error – submit
    else {
      if (event.courseTypeValue === "New" || event.courseNameValue === "New") {
        void newCourseDetails();
      } else {
        void handleSubmit(null);
      }
    }

    setError(error);
  };

  const DeleteEvent = (para: any) => {
    props.setLoaderValue(true);
    SpServices.SPUpdateItem({
      ID: para,
      Listname: "Event",
      RequestJSON: {
        IsDeleted: true,
        isEventCheck: true,
      },
    })
      .then((res: any) => {
        setisDeletePopup({ isOpen: false, Id: null });
        //props.setLoaderValue(false);
        props.setPageConfig({ page: "Dashboard", isToastify: "Delete" });
      })
      .catch((err: any) => {
        ErrFunction("Employee delete error", err);
      });
  };

  const getAttendanceMembers = (event: any) => {
    let membersDetails: any = [];
    SpServices.SPReadItems({
      Listname: "EventAttendance",
      Filter: [
        {
          FilterKey: "EventId",
          Operator: "eq",
          FilterValue: event.ID,
        },
      ],
      Select: "*,Attendees/Title,Attendees/ID,Attendees/EMail",
      Expand: "Attendees",
    })
      .then((_eventAttendance: any) => {
        if (event.Members.length) {
          event.Members.forEach((member: any) => {
            let user = _eventAttendance.filter((val: any) => {
              return val.AttendeesId === member.ID;
            });
            membersDetails.push({
              imgUrl: [...props.SiteUsers].find((_a) => _a.ID === member.ID)
                .imgUrl,
              text: member.Title,
              memberID: member.ID,
              isAttendance: user.length
                ? user[0].Attendant == true
                  ? true
                  : null
                : null,
              eventID: event.ID,
              EventAttendanceID: user.length ? user[0].ID : null,
              isVoluntary: false,
            });
            let SelectAllValue = membersDetails.some(
              (m: any) => m.isAttendance === null
            )
              ? null
              : true;
            setisAttendanceModal({
              isOpen: true,
              Members: event.Members,
              Id: 0,
              selectAll: SelectAllValue,
            });
            setMembersDetails(membersDetails);
          });
        }

        if (event.voluntary.length) {
          event.voluntary.forEach((voluntary: any) => {
            let user = _eventAttendance.filter((val: any) => {
              return val.AttendeesId === voluntary.ID;
            });
            membersDetails.push({
              imgUrl: [...props.SiteUsers].find((_a) => _a.ID === voluntary.ID)
                .imgUrl,
              text: voluntary.Title,
              memberID: voluntary.ID,
              isAttendance: user.length
                ? user[0].Attendant == true
                  ? true
                  : null
                : null,
              eventID: event.ID,
              EventAttendanceID: user.length ? user[0].ID : null,
              isVoluntary: true,
            });

            let SelectAllValue = membersDetails.some(
              (m: any) => m.isAttendance === null
            )
              ? null
              : true;
            setisAttendanceModal({
              isOpen: true,
              Members: event.voluntary,
              Id: 0,
              selectAll: SelectAllValue,
            });
            setMembersDetails(membersDetails);
          });
        }
      })
      .catch((err: any) => console.log("Error get EventAttendance"));
  };

  const handleEventAttendance = () => {
    const _membersDetails = [...membersDetails];
    /*const _members = _membersDetails.filter(
      (member) => member.isAttendance !== null
    );*/

    if (!_membersDetails.length) {
      setisAttendanceModal({
        isOpen: false,
        Members: [],
        Id: 0,
        selectAll: null,
      });
      return;
    }

    props.setLoaderValue(true);

    _membersDetails.forEach((member, i) => {
      const RequestBody = {
        Attendant: member.isAttendance == true ? true : false,
        AttendeesId: member.memberID || null,
        EventIdId: member.eventID || null,
        isVoluntary: member.isVoluntary || false,
      };

      const _index = _membersDetails.findIndex(
        (_a) => _a.memberID === member.memberID
      );

      const handleComplete = (res: any) => {
        if (res?.data?.ID && _index !== -1) {
          _membersDetails[_index]["EventAttendanceID"] = res.data.ID;
          setMembersDetails([..._membersDetails]);
        }

        if (i + 1 === _membersDetails.length) {
          props.setLoaderValue(false);
          setisAttendanceModal({
            isOpen: false,
            Members: [],
            Id: 0,
            selectAll: null,
          });
          props.setPageConfig({ page: "Dashboard", isToastify: "Success" });
        }
      };

      if (member.EventAttendanceID) {
        SpServices.SPUpdateItem({
          Listname: "EventAttendance",
          ID: member.EventAttendanceID,
          RequestJSON: RequestBody,
        })
          .then(handleComplete)
          .catch(() => console.error("Failed to update attendance"));
      } else {
        SpServices.SPAddItem({
          Listname: "EventAttendance",
          RequestJSON: RequestBody,
        })
          .then(handleComplete)
          .catch(() => console.error("Failed to add attendance"));
      }
    });
  };

  const validateAttendance = () => {
    let error = "";

    const hasAttendance = membersDetails.every(
      (member: any) => member.isAttendance === null
    );

    if (hasAttendance) {
      error = "Please mark attendance for at least one person";
    } else {
      handleEventAttendance();
    }

    setError(error);
  };

  // function _onUserChange(items: any[]) {
  //   console.log(items,"Groupitems");
  // }

  // const peoplePickerContext: IPeoplePickerContext = {
  //   absoluteUrl: props.context.pageContext.web.absoluteUrl,
  //   msGraphClientFactory: props.context.msGraphClientFactory,
  //   spHttpClient: props.context.spHttpClient,
  // };

  const voluntary = (event: any) => {
    let _allvoluntary = [...event.voluntaryID];

    const voluntaryDetails = props.SiteUsers.filter((user: any) => {
      return user.secondaryText === props.context.pageContext.user.email;
    });

    _allvoluntary.push(voluntaryDetails[0].ID);
    props.setLoaderValue(true);
    SpServices.SPUpdateItem({
      Listname: "Event",
      ID: Number(event.ID),
      RequestJSON: {
        voluntaryId: {
          results: _allvoluntary ? _allvoluntary : [],
        },
        isEventCheck: true,
      },
    })
      .then((res) => {
        props.setPageConfig({ page: "Dashboard", isToastify: "Success" });
        setPopup(false);
        //props.setLoaderValue(false);
      })
      .catch((err: any) => {
        ErrFunction("Employee Master add data error", err);
      });
  };

  const submitFeedBack = () => {
    let FeedBackJson = {
      Comments: FeedbackText,
      EventIdId: feedBackVisible.ID,
    };

    SpServices.SPAddItem({
      Listname: "feedback",
      RequestJSON: FeedBackJson,
    })
      .then((res) => {
        props.setPageConfig({ page: "Dashboard", isToastify: "Success" });
        setFeedBackVisible({ visible: false, ID: null });
      })
      .catch((err: any) => {
        ErrFunction("feedback add data error", err);
      });
  };

  const feedBackValidation = () => {
    let error = "";
    if (FeedbackText === "") {
      error = "Please enter a feedback";
    } else {
      submitFeedBack();
    }
    setError(error);
  };

  // Function to check if selected date is today
  let EventDate = new Date(EventCurrentValue.Date);
  const isToday =
    EventDate &&
    EventDate.getDate() === today.getDate() &&
    EventDate.getMonth() === today.getMonth() &&
    EventDate.getFullYear() === today.getFullYear();

  // Dynamic minDate for time picker
  const timeMinDate = isToday ? today : undefined;
  React.useEffect(() => {
    filterAndSetEventData(props.EventData, "TodayEvent");
  }, [props.selectedDate]);

  return (
    <>
      <div>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: "15px",
          }}
        >
          <div className="toggleSwitch_container">
            <div className="toggleSwitch_btn_container">
              <i
                style={{
                  position: "absolute",
                  padding: "6px",
                  color: `${
                    activeTab === "TodayEvent" ? "#3f9f4e" : "#6b7280"
                  }`,
                }}
                className="pi pi-calendar"
              ></i>

              <button
                className={`toggleSwitch_btn ${
                  activeTab === "TodayEvent" ? "active" : ""
                }`}
                onClick={() => {
                  setActiveTab("TodayEvent");
                  filterAndSetEventData(props.EventData, "TodayEvent");
                }}
              >
                {/* You can add icon here if needed */}
                Today Event
              </button>
            </div>
            <div className="toggleSwitch_btn_container">
              <i
                style={{
                  position: "absolute",
                  padding: "6px",
                  color: `${activeTab === "MyEvents" ? "#3f9f4e" : "#6b7280"}`,
                }}
                className="pi pi-list"
              ></i>
              <button
                className={`toggleSwitch_btn ${
                  activeTab === "MyEvents" ? "active" : ""
                }`}
                onClick={() => {
                  setActiveTab("MyEvents");
                  filterAndSetEventData(props.EventData, "MyEvents");
                }}
              >
                Upcoming Events
              </button>
            </div>
          </div>

          <div>
            {" "}
            <a
              className={styles.seeAll}
              onClick={() => {
                props.setPageConfig({ page: "Events", isToastify: "" });
              }}
            >
              See all
            </a>
          </div>

          {props.isAdmin &&
            _momentDate.isSameOrAfter(_toDay) &&
            props.SelectedMonth >= todayFormatted && (
              <Button
                className={styles.addButton}
                style={{
                  gap: "4px",
                  fontSize: "12px",
                  position: "absolute",
                  marginLeft: "-126px",
                  marginTop: "9px",
                }}
                icon="pi pi-plus"
                onClick={() => {
                  setPopup(true);
                  setformMode("New");

                  SetEventCurrentValue({
                    ...formData,
                    Date: new Date(
                      `${props.selectedDate.slice(
                        0,
                        4
                      )}-${props.selectedDate.slice(
                        4,
                        6
                      )}-${props.selectedDate.slice(6, 8)}`
                    ),
                  });
                }}
              >
                Add Event
              </Button>
            )}
        </div>

        <div className={styles.formContanier}>
          <div
            style={{ overflow: "auto", height: "340px", marginBottom: "10px" }}
          >
            {curTabEvent?.map((event: any, index: number) => {
              // Optional: Assign a color per index/event type

              const colorIndex = index % colorClasses.length;
              const { bar, color } = colorClasses[colorIndex];

              return (
                <div key={index} className={styles.eventItem}>
                  {/* Left Date Column */}
                  <div className={`${styles.dateColumn} ${bar}`}>
                    <div className={styles.month}>
                      {moment(event.Date).format("MMM").toUpperCase()}
                    </div>
                    <div className={`${styles.day} ${color}`}>
                      {moment(event.Date).format("DD")}
                    </div>
                  </div>

                  {/* Event Details */}
                  <div className={styles.eventContent}>
                    <Tooltip target=".eventTitleTooltip" />
                    <div
                      className={`eventTitleTooltip ${styles.eventTitle}`}
                      data-pr-tooltip={event.Title}
                      data-pr-position="top"
                    >
                      {event.Title}
                    </div>

                    <div className={styles.eventtime}>
                      <i
                        className="pi pi-clock"
                        style={{
                          marginRight: "4px",
                          fontSize: "13px",
                          color: "rgb(102, 102, 102)",
                          paddingTop: "3px",
                        }}
                      />
                      {formatTimeToAMPM(event.StartTime)} -{" "}
                      {formatTimeToAMPM(event.EndTime)}
                    </div>

                    {/* Avatar Row */}
                    <div className={styles.participantsWrapper}>
                      <div className={styles.avatarGroup}>
                        {event.allparticipants
                          ?.slice(0, 5)
                          .map((person: any, i: number) => {
                            const matchedUser = props.SiteUsers.find(
                              (_a: any) => _a.ID === person.ID
                            );
                            // const photoUrl = matchedUser
                            //   ? ${matchedUser.imgUrl}/_layouts/15/userphoto.aspx?size=S&username=${person.EMail}
                            //   : undefined;
                            return (
                              <div key={i} className={styles.avatarCircle}>
                                <img
                                  key={i}
                                  src={matchedUser?.imgUrl}
                                  alt={person.DisplayName}
                                  className={styles.avatar}
                                />
                              </div>
                            );
                          })}
                        {event.allparticipants?.length > 5 && (
                          <span
                            className={styles.extraPeople}
                            data-pr-tooltip={event.allparticipants
                              .slice(5)
                              .map((m: any) => m.DisplayName)
                              .join(", ")}
                          >
                            +{event.allparticipants.length - 5}
                          </span>
                        )}
                      </div>
                    </div>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "flex-end",
                      gap: "6px",
                    }}
                  >
                    {/* Top row: Icons (right-to-left order) */}
                    <div
                      style={{
                        display: "flex",
                        flexDirection: "row-reverse",
                        gap: "10px",
                      }}
                    >
                      {/* Feedback */}
                      {(event.CreatedByEmail ===
                        props.context.pageContext.user.email ||
                        event.allparticipants?.some(
                          (member: any) =>
                            member.EMail ===
                            props.context.pageContext.user.email
                        )) &&
                        isPast(event.Date, event.EndTime) && (
                          <i
                            className="pi pi-comment"
                            title="Feedback"
                            style={{ color: "#56DAE8", cursor: "pointer" }}
                            onClick={() => {
                              void GetEventFeedBack(event);
                              if (
                                event.CreatedByEmail !==
                                props.context.pageContext.user.email
                              ) {
                                setFeedBackVisible({
                                  visible: true,
                                  ID: event.ID,
                                });
                              } else {
                                setUserFeedBackVisible({
                                  visible: true,
                                  ID: event.ID,
                                });
                              }
                            }}
                          />
                        )}

                      {/* Attendance */}
                      {event.CreatedByEmail ===
                        props.context.pageContext.user.email &&
                        props.isAdmin &&
                        isPast(event.Date, event.EndTime) && (
                          <i
                            className="pi pi-users"
                            title="Users"
                            style={{ color: "blue", cursor: "pointer" }}
                            onClick={() => void getAttendanceMembers(event)}
                          />
                        )}

                      {/* Delete */}
                      {event.CreatedByEmail ===
                        props.context.pageContext.user.email &&
                        props.isAdmin &&
                        !isPast(event.Date, event.EndTime) && (
                          <i
                            className="pi pi-trash"
                            title="Delete"
                            style={{ color: "red", cursor: "pointer" }}
                            onClick={() =>
                              setisDeletePopup({ isOpen: true, Id: event.ID })
                            }
                          />
                        )}

                      {/* Edit */}
                      {event.CreatedByEmail ===
                        props.context.pageContext.user.email &&
                        props.isAdmin &&
                        !isPast(event.Date, event.EndTime) && (
                          <i
                          
                            className="pi pi-pencil"
                            title="Edit"
                            style={{ color: "#111f41", cursor: "pointer" }}
                            onClick={() => void getEvent(event.ID, "Edit")}
                          />
                        )}

                      {/* View */}
                      <i
                        className="pi pi-eye"
                        title="View"
                        style={{ color: "#1A8812", cursor: "pointer" }}
                        onClick={() => {
                          void getEvent(event.ID, "View");
                          !isPast(event.Date, event.EndTime);
                        }}
                      />
                    </div>

                    {/* Bottom row: Request button (right aligned) */}
                    <div
                      style={{
                        display: "flex",
                        flexDirection: "row-reverse",
                        width: "100%",
                      }}
                    >
                      {!event.allparticipants?.some(
                        (member: any) =>
                          member.EMail === props.context.pageContext.user.email
                      ) &&
                        props.context.pageContext.user.email !==
                          event.CreatedByEmail &&
                        !isPast(event.Date, event.EndTime) && (
                          <button
                            className={styles.requestBtn}
                            onClick={() => voluntary(event)}
                          >
                            Join
                          </button>
                        )}
                    </div>
                  </div>
                </div>
              );
            })}
            {curTabEvent.length == 0 && (
              <div className={styles.nodatafound}>
                <Label
                  style={{
                    padding: "5px",
                    paddingLeft: "20px",
                    textAlign: "center",
                  }}
                >
                  No events scheduled ...
                </Label>
              </div>
            )}
          </div>
        </div>
      </div>

      <Dialog
        header={formMode === "New" ? "New Event" : "Edit Event"}
        visible={popUp}
        style={{ width: "60vw", maxHeight: "100vh" }}
        onHide={() => {
          setError("");
          setPopup(false);
        }}
        modal
        draggable={false}
        resizable={false}
      >
        <div className={styles.formContainer}>
          {/* Row 1: Title */}
          <div className={styles.formRightContainer}>
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <label className="fieldLabel">Title</label>
                <InputText
                  value={EventCurrentValue.Title}
                  onChange={(data: any) =>
                    handleOnchange("Title", data.target.value)
                  }
                  disabled={formMode == "View"}
                />
              </div>
            </div>
            {/* Row 2: Course Type */}
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <div
                  style={{ display: "flex", justifyContent: "space-between" }}
                >
                  <label className="fieldLabel">Course Type</label>
                  {formMode !== "View" && (
                    <div className="radioGroup">
                      <RadioButton
                        type="radio"
                        name="option"
                        value="Exist"
                        checked={EventCurrentValue.courseTypeValue === "Exist"}
                        onChange={async (e) => {
                          await handleOnchange(
                            "courseTypeValue",
                            e.target.value
                          );
                        }}
                      />
                      <span style={{ padding: "5px" }}>Exist</span>

                      <RadioButton
                        type="radio"
                        name="option"
                        value="New"
                        checked={EventCurrentValue.courseTypeValue === "New"}
                        onChange={async (e) => {
                          await handleOnchange(
                            "courseTypeValue",
                            e.target.value
                          );
                        }}
                      />
                      <span style={{ padding: "5px" }}>New</span>
                    </div>
                  )}
                </div>

                {EventCurrentValue.courseTypeValue === "Exist" ? (
                  <Dropdown
                    options={CourseChoice}
                    value={{
                      name: EventCurrentValue.CourseType,
                      code: EventCurrentValue.CourseType,
                    }}
                    onChange={async (data: any) =>
                      await handleOnchange("CourseType", data.value)
                    }
                    filter
                    filterPlaceholder="Course Type"
                    optionLabel="name"
                    placeholder="Select Course Type"
                    disabled={formMode == "View"}
                  />
                ) : (
                  <InputText
                    value={EventCurrentValue.CourseType}
                    onChange={async (data: any) =>
                      await handleOnchange("CourseType", data.target.value)
                    }
                    disabled={formMode == "View"}
                  />
                )}
              </div>
            </div>
            <div className={styles.formRow}>
              {/* Course Name */}
              <div className={styles.formField}>
                <div
                  style={{ display: "flex", justifyContent: "space-between" }}
                >
                  <label className="fieldLabel">Course Name</label>
                  {formMode !== "View" && (
                    <div className="radioGroup">
                      <RadioButton
                        name="option"
                        value="Exist"
                        checked={EventCurrentValue.courseNameValue === "Exist"}
                        onChange={async (e) => {
                          await handleOnchange(
                            "courseNameValue",
                            e.target.value
                          );
                        }}
                      />
                      <span style={{ padding: "5px" }}>Exist</span>

                      <RadioButton
                        name="option"
                        value="New"
                        checked={EventCurrentValue.courseNameValue === "New"}
                        onChange={async (e) => {
                          await handleOnchange(
                            "courseNameValue",
                            e.target.value
                          );
                        }}
                      />
                      <span style={{ padding: "5px" }}>New</span>
                    </div>
                  )}
                </div>
                {EventCurrentValue.courseNameValue === "Exist" ? (
                  <Dropdown
                    value={{
                      name: EventCurrentValue.CourseName,
                      code: EventCurrentValue.CourseName,
                      ID: EventCurrentValue.CourseConfigId,
                      CourseType: EventCurrentValue.CourseType,
                    }}
                    filter
                    filterPlaceholder="Course Name"
                    optionLabel="name"
                    onChange={(data: any) =>
                      handleOnchange("CourseName", data.value)
                    }
                    placeholder="Select Course Name"
                    options={CourseListChoice.filter(
                      (name: any) =>
                        name.CourseType == EventCurrentValue.CourseType
                    )}
                    disabled={formMode == "View"}
                  />
                ) : (
                  <InputText
                    value={EventCurrentValue.CourseName}
                    onChange={(data: any) =>
                      handleOnchange("CourseName", data.target.value)
                    }
                    disabled={formMode == "View"}
                  />
                )}
              </div>
            </div>
            {/* Description */}
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <label>Description</label>
                <InputTextarea
                  autoResize={false}
                  style={{ resize: "none", overflow: "auto", height: "100px" }}
                  value={EventCurrentValue.Description}
                  onChange={(e) =>
                    handleOnchange("Description", e.target.value)
                  }
                  rows={4}
                  readOnly={formMode === "View"}
                />
              </div>

              <div
                className={styles.formField}
                style={{
                  display: "flex",
                  flexDirection: "row", // side by side
                  gap: "40px", // space between the two blocks
                  alignItems: "center", // vertically align labels and links
                  marginTop: "10px",
                }}
              >
                {/* Join Meeting */}
                {EventCurrentValue?.MeetingLink &&
                  moment
                    .utc()
                    .isBefore(moment.utc(EventCurrentValue.EndTime)) && (
                    <div>
                      <label>Click to join Meeting</label>
                      <br />
                      <a
                        href={EventCurrentValue.MeetingLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ fontWeight: "bold", color: "#00b5ff" }}
                      >
                        Click here
                      </a>
                    </div>
                  )}

                {/* View Meeting Video */}
                {EventCurrentValue?.VideoUrl && (
                  <div>
                    <label>View Meeting Video</label>
                    <br />
                    <a
                      href={`${window.location.origin}${EventCurrentValue.VideoUrl}?web=1`}
                      target="_blank"
                      style={{ fontWeight: "bold", color: "#00b5ff" }}
                      rel="noopener noreferrer"
                    >
                      Click here
                    </a>
                  </div>
                )}
              </div>
            </div>
          </div>
          <div className={`formLeftContainerNew ${styles.formLeftContainer}`}>
            {" "}
            {/* Date & Time */}
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <label className="fieldLabel">Date</label>
                <Calendar
                  value={new Date(EventCurrentValue.Date)}
                  onChange={(e) => handleOnchange("Date", e.value)}
                  showIcon
                  minDate={new Date()}
                  disabled={formMode === "View"}
                />
              </div>
            </div>
            <div
              className={styles.formRow}
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: "15px",
              }}
            >
              <div className={styles.formField}>
                <label className="fieldLabel">Start Time</label>
                <Calendar
                  value={EventCurrentValue.StartTime}
                  onChange={(e) => handleOnchange("StartTime", e.value)}
                  showIcon
                  timeOnly
                  placeholder="Choose start time"
                  disabled={formMode === "View"}
                  minDate={timeMinDate}
                />
              </div>
              <div className={styles.formField}>
                <label className="fieldLabel">End Time</label>
                <Calendar
                  value={EventCurrentValue.EndTime}
                  onChange={(e) => handleOnchange("EndTime", e.value)}
                  showIcon
                  timeOnly
                  disabled={formMode === "View"}
                  minDate={EventCurrentValue.StartTime}
                />
              </div>
            </div>
            {/* Members */}
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <label className="fieldLabel">Members</label>
                <NormalPeoplePicker
                  onResolveSuggestions={GetUserDetails}
                  itemLimit={15}
                  defaultSelectedItems={findUserAll(
                    EventCurrentValue.MembersID
                  )}
                  onChange={(users: any) =>
                    handleOnchange("MembersID", users.length ? users : [])
                  }
                  disabled={formMode === "View"}
                />
              </div>
             
            </div>
            {/* File Upload */}
            <div className={styles.formRow}>
              <div className={styles.UploadBtn}>
                <label htmlFor="fileInput" className={`${styles.uploadlbl}`}>
                  Choose file
                </label>
                <input
                  type="file"
                  id="fileInput"
                  disabled={formMode === "View"}
                  multiple
                  onChange={(e) => getFile(e.target.files)}
                />
                <span id="fileCount">
                  {EventCurrentValue.supportingDocuments &&
                  EventCurrentValue.supportingDocuments.length === 0
                    ? "No files"
                    : `${EventCurrentValue.supportingDocuments.length} file(s)`}
                </span>
              </div>
              <div className={styles.DocContainer}>
                <div style={{ maxHeight: "98px", overflow: "auto" }}>
                  {EventCurrentValue.supportingDocuments.filter(
                    (val: any) => val.Type !== "IsDelete"
                  ).length > 0 ? (
                    <>
                      {EventCurrentValue.supportingDocuments.map(
                        (item: any, index: number) =>
                          item.Types !== "IsDelete" &&
                          (formMode == "New"
                            ? true
                            : item.eventIdId?.includes?.(
                                EventCurrentValue.ID
                              ) === true) && (
                            <div
                              key={index}
                              className="fileItem"
                              style={{
                                padding: "5px 5px",
                                display: "flex",
                                justifyContent: "space-between",
                              }}
                            >
                              <div style={{ display: "flex" }}>
                                <img
                                  style={{ width: "16px", height: "16px" }}
                                  src={getImageUrl(item.name)}
                                  alt=""
                                />
                                <div style={{ padding: "0px 7px" }}>
                                  {item.Types !== "Inlist" ? (
                                    <label>{item.name}</label>
                                  ) : (
                                    <label
                                      style={{
                                        color: "#3f9f4e",
                                        cursor: "pointer",
                                        borderBottom: "1px solid #3f9f4e",
                                      }}
                                      onClick={() => {
                                        window.open(
                                          window.location.origin +
                                            item.url +
                                            "?web=1"
                                        );
                                      }}
                                    >
                                      {item.name}
                                    </label>
                                  )}
                                </div>
                                {formMode !== "View" && (
                                  <div style={{ fontSize: "18px" }}>
                                    <img
                                      style={{
                                        width: "17px",
                                        height: "17px",
                                        padding: "5px 5px 0px 0px",
                                        cursor: "pointer",
                                      }}
                                      src={require("../../components/assets/delete.png")}
                                      onClick={(item: any) => fileDelete(index)}
                                      alt=""
                                    />
                                  </div>
                                )}
                              </div>
                            </div>
                          )
                      )}
                    </>
                  ) : (
                    <>
                      {" "}
                      <div style={{ textAlign: "center", marginTop: "50px" }}>
                        No attachments !!!
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Error */}

        {/* Buttons */}
        <div className={styles.formActions}>
          <div className={styles.errorMsg}>
            <div className="formField fullWidth errorText">
              <b>{Error}</b>
            </div>
          </div>
          <Button
            label="Cancel"
            style={{ background: "#6c757d" }}
            className="p-button-secondary"
            onClick={() => {
              setPopup(false);
              SetEventCurrentValue({ ...formData });
              setError("");
            }}
          />
          {formMode !== "View" && (
            <Button
              label={formMode === "New" ? "Submit" : "Update"}
              onClick={() => Validation(EventCurrentValue)}
            />
          )}
        </div>
      </Dialog>

      <Dialog
        visible={isAttendanceModal.isOpen}
        header={"Event Attendance"}
        style={{ width: "29vw" }}
        // onHide={() => {
        //   setisAttendanceModal({
        //     isOpen: false,
        //     Members: [],
        //     Id: 0,
        //     selectAll: null,
        //   });
        // }}
        draggable={false}
        resizable={false}
        modal
        onHide={() => false}
        closable={false}
        className="addEventAttendance"
      >
        <div className={styles.attendanceBox}>
          {/* <div className={styles.attendanceHeaderBox}>
            <h2 className={styles.attendanceHeader}>Event Attendance</h2>
          </div> */}
          <div className={styles.attendanceContent}>
            <div className={styles.attendanceSelectAllBox}>
              <p className={styles.attendanceSelectAllName}>Select all</p>
              <TriStateCheckbox
                onChange={(e) => {
                  // ✅ Update the selectAll state
                  const newValue = e.value === true ? true : null;

                  setisAttendanceModal({
                    ...isAttendanceModal,
                    selectAll: newValue,
                  });

                  // ✅ Update all members' attendance to match selectAll
                  const updatedMembersDetails = membersDetails.map(
                    (member: any) => ({
                      ...member,
                      isAttendance: newValue,
                    })
                  );

                  setMembersDetails(updatedMembersDetails);
                }}
                /*disabled={
    isEventEnd
      ? true
      : isAllAlreadyAttendant
      ? true
      : isEventStart && !isEventEnd
      ? false
      : true
  }*/
                value={isAttendanceModal.selectAll}
              />
            </div>
            {/* Attendance box */}

            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: "10%",
              }}
            >
              <div style={{ width: "45%" }}>
                <div
                  style={{
                    textAlign: "center",
                    fontSize: "14px",
                    fontWeight: "600",
                    padding: "10px",
                  }}
                >
                  Members
                </div>
                {membersDetails
                  .filter((a: any) => !a.isVoluntary)
                  .map((item: any, index: number) => {
                    return (
                      <div
                        key={index}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "7px",
                          padding: "5px",
                          justifyContent: "space-between",
                          fontSize: "14px",
                        }}
                      >
                        <div
                          className={styles.attendanceImgName}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: "10px",
                          }}
                        >
                          <img
                            src={item.imgUrl}
                            alt="User Image"
                            className={styles.attendancePerMemberImg}
                            onError={(event) => {
                              event.currentTarget.src =
                                "/_layouts/15/images/person.png";
                            }}
                            width={25}
                            height={25}
                            style={{ borderRadius: "25px" }}
                          />
                          <span className={styles.attendanceName}>
                            {item.text}
                          </span>
                        </div>

                        <TriStateCheckbox
                          onChange={(e) => {
                            const _membersDetails = [...membersDetails];
                            const _index = _membersDetails.findIndex(
                              (_a) => _a.memberID === item.memberID
                            );
                            if (_index !== -1) {
                              _membersDetails[_index]["isAttendance"] =
                                e.value === true ? true : null;

                              let SelectAllValue = _membersDetails.some(
                                (m) => m.isAttendance === null
                              )
                                ? null
                                : true;
                              setisAttendanceModal({
                                ...isAttendanceModal,
                                selectAll: SelectAllValue,
                              });
                              setMembersDetails([..._membersDetails]);
                            }
                          }}
                          // disabled={
                          //   isEventEnd
                          //     ? true
                          //     : isAllAlreadyAttendant
                          //     ? true
                          //     : isEventStart && !isEventEnd
                          //     ? false
                          //     : true
                          // }

                          value={item.isAttendance == true ? true : null}
                        ></TriStateCheckbox>
                      </div>
                    );
                  })}
              </div>
              <div style={{ width: "45%" }}>
                <div
                  style={{
                    textAlign: "center",
                    fontSize: "14px",
                    fontWeight: "600",
                    padding: "10px",
                  }}
                >
                  Voluntary
                </div>
                {membersDetails
                  .filter((a: any) => a.isVoluntary)
                  .map((item: any, index: number) => {
                    return (
                      <div
                        key={index}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "7px",
                          padding: "5px",
                          justifyContent: "space-between",
                          fontSize: "14px",
                        }}
                      >
                        <div
                          className={styles.attendanceImgName}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: "10px",
                          }}
                        >
                          <img
                            src={item.imgUrl}
                            alt="User Image"
                            className={styles.attendancePerMemberImg}
                            onError={(event) => {
                              event.currentTarget.src =
                                "/_layouts/15/images/person.png";
                            }}
                            width={25}
                            height={25}
                            style={{ borderRadius: "25px" }}
                          />
                          <span className={styles.attendanceName}>
                            {item.text}
                          </span>
                        </div>

                        <TriStateCheckbox
                          onChange={(e) => {
                            const _membersDetails = [...membersDetails];
                            const _index = _membersDetails.findIndex(
                              (_a) => _a.memberID === item.memberID
                            );
                            if (_index !== -1) {
                              _membersDetails[_index]["isAttendance"] =
                                e.value == true ? true : null;

                              let SelectAllValue = _membersDetails.some(
                                (m) => m.isAttendance === null
                              )
                                ? null
                                : true;

                              setisAttendanceModal({
                                ...isAttendanceModal,
                                selectAll: SelectAllValue,
                              });
                              setMembersDetails([..._membersDetails]);
                            }
                          }}
                          // disabled={
                          //   isEventEnd
                          //     ? true
                          //     : isAllAlreadyAttendant
                          //     ? true
                          //     : isEventStart && !isEventEnd
                          //     ? false
                          //     : true
                          // }

                          value={item.isAttendance == true ? true : null}
                        ></TriStateCheckbox>
                      </div>
                    );
                  })}
              </div>
            </div>
          </div>
          <div className={styles.errorMsg} style={{ textAlign: "center" }}>
            <div className="formField fullWidth errorText">
              <b>{Error}</b>
            </div>
          </div>

          <div className={styles.attendanceFooterBox}>
            <Button
              className={styles.cancelBtn}
              style={{
                height: "25px",
                width: "65px",
                fontWeight: "400",
                fontSize: "14px",
              }}
              onClick={(data: any) => {
                setisAttendanceModal({
                  isOpen: false,
                  Members: [],
                  Id: 0,
                  selectAll: null,
                });

                setError("");
              }}
            >
              Close{" "}
            </Button>
            <Button
              className={styles.addBtn}
              style={{
                height: "25px",
                width: "65px",
                fontWeight: "400",
                fontSize: "14px",
              }}
              onClick={(data: any) => {
                void validateAttendance();
              }}
            >
              Submit
            </Button>
          </div>
        </div>
      </Dialog>
      {/* FeedBack PopUp*/}
      <Dialog
        visible={feedBackVisible.visible}
        draggable={false}
        resizable={false}
        modal
        style={{ width: "25rem" }}
        onHide={() => setFeedBackVisible({ visible: false, ID: null })}
        header="Your Feedback!"
        className="feedbackDialog"
      >
        <div className={styles.formField}>
          <InputTextarea
            value={FeedbackText || ""}
            autoResize={false}
            placeholder="Enter your feedback here..."
            style={{ resize: "none", overflow: "auto", height: "100px" }}
            onChange={(e: any) => setFeedbackText(e.target.value)}
            disabled={!!currentFeedback}
          />
        </div>
        <div className={styles.errorMsg} style={{ textAlign: "center" }}>
          <div className="formField fullWidth errorText">
            <b>{Error}</b>
          </div>
        </div>

        {!currentFeedback && (
          <div
            style={{
              padding: "20px 0px 0px 0px",
              display: "flex",
              gap: "11px",
              justifyContent: "center",
            }}
          >
            <Button
              style={{ height: "30px", width: "80px" }}
              className={styles.addBtn}
              onClick={() => {
                feedBackValidation();
              }}
            >
              Submit
            </Button>

            <Button
              style={{ height: "30px", width: "80px" }}
              className={styles.cancelBtn}
              onClick={() => {
                setFeedBackVisible({ visible: false, ID: null });
                setError("");
                //
              }}
            >
              Cancel
            </Button>
          </div>
        )}
      </Dialog>
      <Dialog
        visible={UserfeedBackVisible.visible}
        modal
        draggable={false}
        resizable={false}
        style={{ width: "40rem" }}
        onHide={() => setUserFeedBackVisible({ visible: false, ID: null })}
        header="Meeting Feedback"
        className="UserfeedbackDialog"
      >
        <div className={styles.cardGrid}>
          {feedbackList.length === 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                height: "200px",
                width: "100%",
              }}
            >
              <Label
                style={{
                  padding: "10px 20px",
                  textAlign: "center",
                  fontSize: "16px",
                  fontWeight: "500",
                }}
              >
                No feedback found...
              </Label>
            </div>
          ) : (
            feedbackList.map((fb: any, i: number) => (
              <div key={i} className={styles.card}>
                <p className={styles.comment}>"{fb.comments}"</p>
                <div
                  style={{ display: "flex", alignItems: "center", gap: "10px" }}
                >
                  <img
                    src={`/_layouts/15/userphoto.aspx?size=L&username=${fb.CreatedByEmail}`}
                    className={styles.cardAvatar}
                    alt={fb.CreatedBy}
                  />
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "flex-start",
                    }}
                  >
                    <h4 className={styles.userName}>{fb.CreatedBy}</h4>
                    <span style={{ fontSize: "11px" }}>
                      {moment(fb.CreatedDate).fromNow()}
                    </span>
                  </div>
                </div>
              </div>
            ))
          )}
        </div>
      </Dialog>

      <Dialog
        visible={isDeletePopup.isOpen}
        modal
        style={{ width: "25rem" }}
        onHide={() => setisDeletePopup({ isOpen: false, Id: null })}
        header="Confirmation"
        className="deleteDialog"
      >
        <div style={{ padding: "20px", textAlign: "center", fontSize: "14px" }}>
          Are you sure you want to delete this event?
        </div>
        <div
          style={{
            display: "flex",
            gap: "11px",
            justifyContent: "center",
          }}
        >
          <Button
            style={{
              height: "25px",
              width: "65px",
              background: "#aa1f1f",
              borderRadius: "3px",
            }}
            onClick={() => {
              DeleteEvent(isDeletePopup.Id);
            }}
          >
            Yes
          </Button>

          <Button
            style={{ height: "25px", width: "65px", borderRadius: "3px" }}
            className={styles.cancelBtn}
            onClick={() => {
              setisDeletePopup({ isOpen: false, Id: null });
              //
            }}
          >
            No
          </Button>
        </div>
      </Dialog>
    </>
  );
};

export default EventDetails;
