import { sp } from "@pnp/pnpjs";
import * as React from "react";
import { useEffect, useState, useRef } from "react";
//import { Carousel } from "primereact/carousel";
import styles from "./Banner.module.scss";
import Calendarview from "../Calander/Calander";
import SpServices from "../../../../CommonServices/SpServices";
import EventDetails from "../UpcomingEvents/UpcomingEvents";
import * as moment from "moment";
import Loader from "../Loader/Loader";
import "../../../../ExternalRef/CSS/Style.css";
import Certification from "../Certification/Certification";
import Documents from "../Document/Documents";
import Contributor from "../Contributor/Contributor";
import VideoComponent from "../Video/Video";
import EventsDashboard from "../EventsDashboard/EventsDashboard";
import DocumentDashBoard from "../DocumentDashBoard/DocumentDashBoard";
import { IPageConfig } from "../../Config/Interface";
import { IEventDetails } from "../../Config/Interface";
import CertDashboard from "../CertDashboard/CertDashboard";
import { Toast } from "primereact/toast";
let _sDate: string = moment(new Date()).format("YYYYMMDD");

const Banner = (props: any) => {
  const _pageConfig: IPageConfig = {
    page: "Dashboard",
    isToastify: "",
  };
  const [pageConfig, setPageConfig] = useState<IPageConfig>({ ..._pageConfig });
  const [EventData, setEventData] = useState<IEventDetails.IEventInfo[]>([]);
  const [overAllRefresh, setOverAllRefresh] = useState(false);
  const [LoaderValue, setLoaderValue] = useState(false);
  const [selectedDate, setSelectedDate] = React.useState(
    moment(new Date()).format("YYYYMMDD")
  );
  const [SiteUsers, setSiteUsers] = useState<any>([]);
  const [userRole, setUserRole] = useState<string>("");
  const [documents, setDocuments] = useState<IEventDetails.IDocuments[]>([]);
  const [isVideo, setIsvisible] = useState(false);
  const overAllRefreshFunction = () => {
    setOverAllRefresh((prv) => !prv);
  };
  const [videoDetails, setVideoDetails] = React.useState<
    IEventDetails.IVideo[]
  >([]);
  const [recentCertifications, setRecentCertifications] = React.useState([]);
  const [Certifications, setCertifications] = React.useState([]);
  const [SelectedMonth, setSelectedMonth] = useState("");
  //const [testSelecteddate,settestSelecteddate] = useState(moment(new Date()).format("YYYYMMDD"))

  const ErrFunction = (err: string, funcName: string): void => {
    console.error(funcName, err);
  };

  const toast = useRef<Toast>(null);

  const showSuccess = (severity: any, summary: any, detail: any) => {
    toast.current?.show({
      severity: severity,
      summary: summary,
      detail: detail,
      life: 3000,
    });
  };
let SiteName=props.context._pageContext._site.absoluteUrl.split('com')[1];
  function _sortFilterKeys(a: any, b: any) {
    if (a.TimeCreated > b.TimeCreated) {
      return -1;
    }
    if (a.TimeCreated < b.TimeCreated) {
      return 1;
    }
    return 0;
  }
  const collectCertificationsDetails = async () => {
    try {
      const items: any = await sp.web.lists
        .getByTitle("CertificationsDetails")
        .items.select("*,EmployeeName/Title,EmployeeName/ID,EmployeeName/EMail")
        .expand("EmployeeName")
        .orderBy("Date", false)
        .top(999)
        .get();

      const allCertifications = items.map((item: any) => ({
        EmployeeName: item.EmployeeName?.Title || "",
        EmployeeNameID: item.EmployeeName?.ID || null,
        EmployeeNameEmail: item.EmployeeName?.EMail || "",
        CertificationName: item.CertificationName || "",
        Department: item.Department || "",
        Date: item.Date,
      }));

      // Filter latest per employee
      const latestPerEmployee: any = [];
      for (let cert of allCertifications) {
        if (
          !latestPerEmployee.some(
            (e: any) => e.EmployeeNameID === cert.EmployeeNameID
          )
        ) {
          latestPerEmployee.push(cert);
        }
      }

      if (pageConfig.isToastify == "Success") {
        showSuccess("success", "Success", "Successfully added your data");
      } else if (pageConfig.isToastify == "Delete") {
        showSuccess("warn", "Warning", "Successfully deleted your record");
      } else if (pageConfig.isToastify == "Info") {
        showSuccess("info", "Info", "Successfully updated your record");
      }

      setRecentCertifications(latestPerEmployee.slice(0, 6));
      setCertifications(allCertifications);
      setLoaderValue(false);
    } catch (err) {
      console.error("Get CertificationDetails error", err);
    }
  };

  const getVideo = async () => {
    try {
      const files = await sp.web
        .getFolderByServerRelativeUrl(`${SiteName}/RecordingVideos`)
        .files.select("Name,ServerRelativeUrl,TimeCreated")
        .get();
      console.log(files, "Fetched files from SharePoint");
      if (files.length > 0) {
        const _video: any[] = files.map((file: any) => ({
          name: file.Name,
          url: file.ServerRelativeUrl,
          type: file.Name.split(".").pop() || "mp4",
          TimeCreated: file.TimeCreated,
        }));
        setVideoDetails([..._video].sort(_sortFilterKeys));
        void collectCertificationsDetails();
      } else {
        void collectCertificationsDetails();
      }
    } catch (error) {
      console.error("Error fetching videos:", error);
    }
  };

 const getAllDocuments = async () => {
  await SpServices.SPReadItems({
    Listname: "Documents",
    Select:
      "*,File/Name,File/ServerRelativeUrl,Author/Title,Author/ID,Author/EMail,Folder/Name,Folder/ServerRelativeUrl,FSObjType",
    Expand: "File,Author,Folder",
    Orderby: "Modified",
    Orderbydecorasc: false,
    Topcount: 999,
  })
    .then(async (docs: any) => {
      let _folders = docs.filter(
        (_a: any) => _a.FSObjType === 1 && _a.FileSystemObjectType === 1
      );
      let _files = docs.filter((_a: any) => _a.FSObjType === 0);
      let _documents: any = [];

      if (_folders.length) {
        for (const _folder of _folders) {
          const folderUrl = _folder.Folder.ServerRelativeUrl;

          // match only direct files inside folder
          const relatedFiles = _files.filter((file: any) => {
            const fileUrl = file.File.ServerRelativeUrl;
            const parentUrl = folderUrl + "/";
            const relativePart = fileUrl.replace(parentUrl, "");
            return relativePart && !relativePart.includes("/");
          });

          // map file details
          const _attachments = relatedFiles.map((_docs: any) => ({
            name: _docs.File.Name,
            url: _docs.File.ServerRelativeUrl,
            type: _docs.File.Name.split(".").pop(),
            time: _docs.Modified,
            CreatedByName: _docs.Author?.Title || "",
          }));

          // push folder-level info (Added time & CreatedByName)
          _documents.push({
            FolderName: _folder.Folder.Name,
            FolderUrl: _folder.Folder.ServerRelativeUrl,
            time: _folder.Modified, // folder modified time
            CreatedByName: _folder.Author?.Title || "", // folder creator/author
            supportingDocuments: [..._attachments],
          });
        }

        console.log(_documents, "docs");
        setDocuments(_documents.sort(_sortFilterKeys));
        void getVideo();
      } else {
        void getVideo();
      }
    })
    .catch((err: any) => {
      ErrFunction("Event Attachment Data get error", err);
    });
};


  const collectEventDetailsData = async () => {
    await sp.web.lists
      .getByTitle("Event")
      .items.select(
        "*,Members/Title,Members/ID,Members/EMail,voluntary/Title,voluntary/ID,voluntary/EMail,Author/Title,Author/ID,Author/EMail"
      )
      .expand("Members,voluntary,Author")
      .filter("IsDeleted eq 0")
      .get()
      .then((events) => {
        const eventsDetails: IEventDetails.IEventInfo[] = [];
        // let _attachement: IFiles[] = [];

        if (events.length) {
          let completed = 0;

          for (let i = 0; i < events.length; i++) {
            const event = events[i];
            void SpServices.SPReadItemUsingId({
              Listname: "CourseConfig",
              SelectedId: event?.CourseConfigId || null,
              Select: "*",
              Expand: "",
            })
              .then(async (CourseItem: any) => {
                eventsDetails.push({
                  CourseType: CourseItem?.CourseType || "",
                  CourseName: CourseItem?.CourseName || "",
                  Title: event.Title || "",
                  Date: new Date(event.Date) || new Date(),
                  Members: event.Members?.length ? event.Members : [],
                  MembersID: event.MembersId?.length ? event.MembersId : [],
                  IsDeleted: false,
                  Description: event.Description || "",
                  EndTime: event.EndTime || null,
                  StartTime: event.StartTime || null,
                  CourseConfigId: event.CourseConfigId || null,
                  ID: event.ID || null,
                  supportingDocuments: [],
                  FolderName: "",
                  voluntary: event.voluntary?.length ? event.voluntary : [],
                  voluntaryID: event.voluntaryId?.length
                    ? event.voluntaryId
                    : [],
                  CreatedBy: event.Author?.Title || "",
                  CreatedByID: event.Author?.ID || null,
                  CreatedByEmail: event.Author?.EMail || "",
                  modifiedon: event.Modified || "",
                  isEventCheck: false,
                  allparticipants: [
                    ...(event.voluntary?.length ? event.voluntary : []),
                    ...(event.Members?.length ? event.Members : []),
                  ],
                });

                completed++;

                if (completed === events.length) {
                  console.log(eventsDetails, "ev");
                  setEventData(eventsDetails);
                  void getAllDocuments();
                }
              })
              .catch((err: any) => {
                ErrFunction("Event  Data get error", err);
              });
          }
        } else {
          void getAllDocuments();
        }
      })
      .catch((err) => {
        console.log("Get EventDetails error", err);
      });
  };




  
  const getSiteUsers = () => {
    SpServices.getAllUsers()
      .then((res: any[]) => {
        let sitePeople: any = [];
        let count = 0;
        if (res.length) {
          res.forEach(async (user: any) => {
            count++;
            if (user.Email !== "") {
              sitePeople.push({
                key: 1,
                imgUrl: `/_layouts/15/userphoto.aspx?size=S&accountname= +
                  ${user.Email}`,
                text: user.Title,
                ID: user.Id,
                secondaryText: user.Email,
                isValid: true,
              });
            }
            if (user.Email === props.context.pageContext.user.email) {
              let userProfile = "";
              userProfile = await sp.profiles.getUserProfilePropertyFor(
                user.LoginName,
                "Title"
              );
              setUserRole(userProfile || "");
            }

            if (count == res.length) {
              setSiteUsers([...sitePeople]);

              void collectEventDetailsData();
            }
          });
        } else {
          void collectEventDetailsData();
        }
      })
      .catch((err) => console.log("Get site users", err));
  };

  useEffect(() => {
    console.log(props.context._pageContext._site.absoluteUrl.split('com')[1],"Site")
    if (pageConfig.page == "Dashboard") {
      setLoaderValue(true);
      void getSiteUsers();
    }
  }, [pageConfig]);

  return LoaderValue ? (
    <Loader />
  ) : (
    <>
      <Toast ref={toast} />
      <div className={styles.OverAllDiv}>
        <div className={styles.companyBannerContainer}>
          <div className={styles.banner_container}>
            <div className={styles.banner_center}>
              <h2
                style={{
                  color: "#3f9f4e",
                  fontSize: "19px",
                  fontWeight: 700,
                  margin: 0,
                }}
              >
                Learning Management System
              </h2>
            </div>
            <div className={styles.banner_right}>
              <div className={styles.banner_right_text}>
                <strong>
                  Welcome, {props.context.pageContext.user.displayName}
                </strong>

                <div className={styles.role}>
                  <i
                    className="pi pi-shield"
                    style={{ fontSize: "13px", color: "#3f9f4e" }}
                  ></i>
                  <span>{userRole}</span>
                </div>
              </div>
            </div>
          </div>
        </div>
        {pageConfig.page == "Events" ? (
          <EventsDashboard
            SiteUsers={SiteUsers}
            setPageConfig={setPageConfig}
          />
        ) : pageConfig.page == "Documents" ? (
          <DocumentDashBoard
            Document={documents}
            videoDetails={videoDetails}
            setPageConfig={setPageConfig}
            context={props.context}
            isVideo={isVideo}
          />
        ) : pageConfig.page == "Certificates" ? (
          <CertDashboard
            setPageConfig={setPageConfig}
            Certifications={Certifications}
            SiteUsers={SiteUsers}
          />
        ) : (
          <>
            <div className={styles.BannerImg}>
              <div className={styles.BannerContent}>
                "Every click, Every lesson, Every challenge you complete is an
                investment in the most valuable asset you own"
              </div>
            </div>

            <div className={styles.RowContainer}>
              <div className={styles.CalendarEventWrapper}>
                <div className={styles.CalendarEventContainer}>
                  <div className={styles.firstHalfWidth}>
                    <Calendarview
                      setSelectedDate={setSelectedDate}
                      selectedDate={selectedDate}
                      setSelectedMonth={setSelectedMonth}
                      _sDate={_sDate}
                    />
                  </div>
                  <div className={styles.HalfWidth}>
                    <EventDetails
                      selectedDate={selectedDate}
                      LoaderValue={LoaderValue}
                      setLoaderValue={setLoaderValue}
                      SiteUsers={SiteUsers}
                      context={props.context}
                      isAdmin={props.isAdmin}
                      setIsAdmin={props.setIsAdmin}
                      overAllRefresh={overAllRefresh}
                      overAllRefreshFunction={overAllRefreshFunction}
                      setPageConfig={setPageConfig}
                      SelectedMonth={SelectedMonth}
                      EventData={EventData}
                      showSuccess={showSuccess}
                      _sDate={_sDate}
                      SiteName={SiteName}
                    />
                  </div>
                </div>
              </div>
              <div className={styles.ContributorContainer}>
                <Contributor EventData={EventData} SiteUsers={SiteUsers} />
              </div>
            </div>

            <div
              className={styles.RowContainer}
              style={{ backgroundColor: "white", height: "505px" }}
            >
              <div className={styles.DocumentVideoWrapper}>
                <div className={styles.DocumentVideoContainer}>
                  <div className={styles.HalfWidth}>
                    <VideoComponent
                      videoDetails={videoDetails}
                      setPageConfig={setPageConfig}
                      setIsvisible={setIsvisible}
                    />
                  </div>
                  <div className={styles.HalfWidth}>
                    <Documents
                      Document={documents}
                      setPageConfig={setPageConfig}
                      setIsvisible={setIsvisible}
                      context={props.context}
                    />
                  </div>
                </div>
              </div>
              <div className={styles.CertificationContainer}>
                <Certification
                  SiteUsers={SiteUsers}
                  setPageConfig={setPageConfig}
                  recentCertifications={recentCertifications}
                />
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
};

export default Banner;
