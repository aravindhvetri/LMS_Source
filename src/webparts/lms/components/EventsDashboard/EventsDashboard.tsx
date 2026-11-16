import * as React from "react";
import SpServices from "../../../../CommonServices/SpServices";
import styles from "./EventsDashboard.module.scss";
import * as moment from "moment";

import { Label } from "@fluentui/react/lib/Label";
import { Tooltip } from "primereact/tooltip";
import { InputText } from "primereact/inputtext";

function EventsDashboard(props: any) {
  const [EventDetails, setEventDetails] = React.useState<any>([]);
  const [filteredEvents, setFilteredEvents] = React.useState<any>([]);
  const [searchTerm, setSearchTerm] = React.useState<string>("");

  const ErrFunction = (err: string, funcName: string): void => {
    console.error(funcName, err);
  };

  const fetchEvents = async () => {
    try {
      const events: any = await SpServices.SPReadItems({
        Listname: "Event",
        Select:
          "*,Members/Title,Members/ID,Members/EMail,voluntary/Title,voluntary/ID,voluntary/EMail,Author/Title,Author/ID,Author/EMail",
        Expand: "Members,voluntary,Author",
        Orderby: "Modified",
        Filter: [
          {
            FilterKey: "IsDeleted",
            Operator: "eq",
            FilterValue: "0",
          },
        ],
      });

      let _EventData: any = [];

      for (const event of events) {
        try {
          const CourseItem: any = await SpServices.SPReadItemUsingId({
            Listname: "CourseConfig",
            SelectedId: event?.CourseConfigId || null,
            Select: "*",
            Expand: "",
          });

          _EventData.push({
            CourseType: CourseItem.CourseType || "",
            CourseName: CourseItem.CourseName || "",
            Title: event.Title || "",
            Date: event.Date || new Date(),
            Members: event.Members?.length ? event.Members : [],
            MembersID: event.MembersId?.length ? event.MembersId : [],
            IsDeleted: false,
            Description: event.Description || "",
            EndTime: event.EndTime || "",
            StartTime: event.StartTime || "",
            CourseConfigId: event.CourseConfigId || null,
            ID: event.ID || null,
            voluntary: event.voluntary?.length ? event.voluntary : [],
            voluntaryID: event.voluntaryId?.length ? event.voluntaryId : [],
            CreatedBy: event.Author?.Title || "",
            CreatedByID: event.Author?.ID || null,
            CreatedByEmail: event.Author?.EMail || "",
            modifiedon: event.Modified,
            allparticipants: [
              ...(event.voluntary?.length ? event.voluntary : []),
              ...(event.Members?.length ? event.Members : []),
            ],
            MeetingLink:event.MeetingLink
          });
        } catch (err) {
          ErrFunction("CourseConfig fetch error", "SPReadItemUsingId");
        }
      }

      setEventDetails(_EventData);
      setFilteredEvents(_EventData);
    } catch (err) {
      ErrFunction("Error fetching events", "fetchEvents");
    }
  };

  function formatTimeToAMPM(timeStr: any) {
    if (!timeStr) return "";
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

  const colorClasses = [
    { bar: styles.tealBar, color: styles.tealcolor },
    { bar: styles.yellowBar, color: styles.yellowcolor },
    { bar: styles.greenBar, color: styles.greencolor },
    { bar: styles.purpleBar, color: styles.purplecolor },
  ];

  const handleSearch = (value: string) => {
    setSearchTerm(value);
    if (!value.trim()) {
      setFilteredEvents(EventDetails);
      return;
    }

    const searchValue = value.toLowerCase().replace(/\s+/g, "");
    const filtered = EventDetails.filter((event: any) => {
      return event.Title?.toLowerCase()
        .replace(/\s+/g, "")
        .includes(searchValue);
    });
    setFilteredEvents(filtered);
  };

  React.useEffect(() => {
    void fetchEvents();
  }, []);

  return (
    <>
      {/* Back + Heading */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "10px",
          marginBottom: "20px",
          justifyContent: "space-between",
        }}
      >
        <div
          onClick={() => {
            props.setPageConfig({ page: "Dashboard" });
          }}
          
        >
          <div style={{display:"flex",gap:"10px"}}>
            <i
              className="pi pi-arrow-left"
              style={{ marginTop: "4px", cursor: "pointer" }}
            ></i>
            <h4 className={styles.heading}>All Events</h4>
          </div>
        </div>

        {/* Search Box */}
        <div style={{ marginBottom: "20px" }}>
          <span className="p-input-icon-left" style={{ width: "300px" }}>
            <i className="pi pi-search" style={{ paddingLeft: "8px",height:"10x", color:"green"}} />
            <InputText
              value={searchTerm}
              onChange={(e) => handleSearch(e.target.value)}
              placeholder="Search by Event"
              style={{ width: "100%", padding: "10px 30px" }}
            />
          </span>
        </div>
      </div>

      {/* Events List */}
      <div className={styles.Dashboard_container}>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "1%",
            marginBottom: "20px",
          }}
        >
          {filteredEvents?.map((event: any, index: number) => {
            const colorIndex = index % colorClasses.length;
            const { bar, color } = colorClasses[colorIndex];

            return (
              <div key={index} className={styles.eventItem}>
                {/* Date Column */}
                <div className={`${styles.dateColumn} ${bar}`}>
                  <div className={styles.month}>
                    {moment(event.Date).format("MMM").toUpperCase()}
                  </div>
                  <div className={`${styles.day} ${color}`}>
                    {moment(event.Date).format("DD")}
                  </div>
                </div>

                {/* Event Content */}
                <div className={styles.eventContent}>
                  <Tooltip target=".eventTitleTooltip" />
                  <div
                    className={`eventTitleTooltip ${styles.eventTitle}`}
                    data-pr-tooltip={event.Title}
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

                  {/* Participants */}
                  <div className={styles.avatarGroup}>
                    {event.allparticipants
                      ?.slice(0, 5)
                      .map((person: any, i: number) => {
                        const matchedUser = props.SiteUsers.find(
                          (_a: any) => _a.ID === person.ID
                        );
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
            );
          })}

          {filteredEvents.length === 0 && (
            <div className={styles.nodatafound}>
              <Label>No events found ...</Label>
            </div>
          )}
        </div>
      </div>
    </>
  );
}

export default EventsDashboard;
