import * as React from "react";
import { useEffect } from "react";
import styles from "./Calander.module.scss";

import { Calendar } from "@fullcalendar/core";
import interactionPlugin from "@fullcalendar/interaction";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import listPlugin from "@fullcalendar/list";
import bootstrap5Plugin from "@fullcalendar/bootstrap5";
import * as moment from "moment";

const Calendarview = (props: any) => {
  const BindCalender = () => {
    let calendarEl: any = document.getElementById("myCalendar");

    let _Calendar = new Calendar(calendarEl, {
      plugins: [
        interactionPlugin,
        dayGridPlugin,
        timeGridPlugin,
        listPlugin,
        bootstrap5Plugin,
      ],
      selectable: true,
      buttonText: {
        today: "Today",
        dayGridMonth: "Month",
        dayGridWeek: "Week - ListGrid",
        timeGridWeek: "Week",
        dayGridDay: "Day - ListGrid",
        timeGridDay: "Day",
        listWeek: "List",
      },
      headerToolbar: {
        left: "prev next",
        center: "title",
        right: "",
      },
      initialDate: new Date(),
      events: [],
      height: "auto",
      displayEventTime: true,
      weekends: true,
      dayMaxEventRows: true,
      views: {
        dayGrid: {
          dayMaxEventRows: 4,
        },
      },
      showNonCurrentDates: false,
      fixedWeekCount: false,

      eventDidMount: (event) => {
        event.el.setAttribute("data-bs-target", "event");
      },

      dayCellDidMount: function (info) {
        const today = moment().startOf("day");
        const cellDate = moment(info.date).startOf("day");

        // Highlight selected date dynamically
        if (
          props.selectedDate &&
          cellDate.isSame(moment(props.selectedDate, "YYYYMMDD"))
        ) {
          document
            .querySelectorAll(`.${styles.fcSelectedDate}`)
            .forEach((el) => {
              el.classList.remove(styles.fcSelectedDate);
            });
          info.el.classList.add(styles.testDate);
        }

        // Highlight today
        if (cellDate.isSame(today)) {
          info.el.classList.add(styles.fcTodayCustom);
        }

        // Highlight first 7 days of week
        const startOfWeek = today.clone().startOf("week");
        const endOfWeek = today.clone().endOf("week");
        if (cellDate.isBetween(startOfWeek, endOfWeek, null, "[]")) {
          info.el.classList.add(styles.fcFirst7Days);
        }
      },

      dateClick: function (arg) {
        document.querySelectorAll(`.${styles.fcSelectedDate}`).forEach((el) => {
          el.classList.remove(styles.fcSelectedDate);
        });
        document
            .querySelectorAll(`.${styles.testDate}`)
            .forEach((el) => {
              el.classList.remove(styles.testDate);
            });

        arg.dayEl.classList.add(styles.fcSelectedDate);
        const selectedDateString = moment(arg.dateStr).format("YYYYMMDD");
        props._sDate = selectedDateString;
        props.setSelectedDate(
          selectedDateString || moment().format("YYYYMMDD")
        );
      },
      datesSet: function (info) {
        const year = info.start.getFullYear();
        const month = String(info.start.getMonth() + 1).padStart(2, "0");
        const currentMonthFormatted = `${year}${month}`;
        props.setSelectedMonth(currentMonthFormatted);

        if (props._sDate) {
          // Format once
          const selectedDateFormatted = moment(
            // props.selectedDate,
            props._sDate,
            "YYYYMMDD"
          ).format("YYYY-MM-DD");

          // Remove old highlights
          document
            .querySelectorAll(`.${styles.fcSelectedDate}`)
            .forEach((el) => {
              el.classList.remove(styles.fcSelectedDate);
            });

          // Delay to ensure cells are rendered
          setTimeout(() => {
            const selectedDateEl = document.querySelector(
              `[data-date="${selectedDateFormatted}"]`
            );

            if (selectedDateEl) {
              selectedDateEl.classList.add(styles.fcSelectedDate);
            }
          }, 0);
        }
      },
    });

    _Calendar.updateSize();
    _Calendar.render();
  };

  useEffect(() => {
    void BindCalender();
    console.log(props.selectedDate, "props.selectedDateCalander");
  }, []);

  return (
    <div className={styles.calendercontainer}>
      <div className={styles.calenderEventHeader}>
        <h4 className={styles.Title}>Events Calendar</h4>
      </div>

      <div className={styles.CalCardContainer}>
        <div style={{ width: "100%" }}>
          <div id="myCalendar"></div>
        </div>
      </div>
    </div>
  );
};

export default Calendarview;
