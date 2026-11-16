import * as React from "react";
import "../../../../ExternalRef/CSS/Style.css";

import styles from "./Certification.module.scss";
import { Avatar } from "primereact/avatar";
import * as moment from "moment";
import { Label } from "@fluentui/react";
import { Tooltip } from "primereact/tooltip";


interface ICertificationInfo {
  EmployeeName: string;
  EmployeeNameID: number | null;
  CertificationName: string;
  Department: string;
  Date: any;
  EmployeeNameEmail: string;
}

const Certification = (props: any) => {
  const [recentCertifications, setRecentCertifications] = React.useState<
    ICertificationInfo[]
  >([]);

  React.useEffect(() => {
    setRecentCertifications(props.recentCertifications);
  }, [props.recentCertifications]);

  return (
    <div>
      <div className={styles.header}>
        <h4 className={styles.title}>Certification Achievements</h4>

        {recentCertifications.length > 0 && (
          <a
            className={styles.seeAll}
            onClick={() => {
              props.setPageConfig({ page: "Certificates" });
            }}
          >
            See all
          </a>
        )}
      </div>
      {recentCertifications.length > 0 ? (
        recentCertifications.map((cert, index) => {
          const matchedUser = props.SiteUsers?.find(
            (user: any) => user.ID === cert.EmployeeNameID
          );

          return (
            <div className={styles.certCard} key={index}>
              <Avatar
                image={matchedUser?.imgUrl}
                shape="circle"
                size="large"
                className={styles.avatar}
              />
              <div className={styles.userInfo}>
                <div className={styles.nameRow}>
                  <Tooltip target=".eventTitleTooltip" />
                  <div
                    className={`eventTitleTooltip ${styles.eventTitle}`}
                    data-pr-tooltip={cert.EmployeeName}
                     data-pr-position="top"
                  >
                    {cert.EmployeeName}
                  </div>
                 
                </div>
                <div className={styles.subtitle}>
                  <Tooltip target=".certName" />
                  <div
                    className={`certName ${styles.eventCertificationName}`}
                    data-pr-tooltip={cert.CertificationName}
                     data-pr-position="top"
                  >
                    {cert.CertificationName}
                  </div>
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  flexDirection: "column",
                }}
              >
                <img
                  src={require("../../components/assets/mark.png")}
                  width="30"
                  height="30"
                  alt="Mark"
                />
                <div className={styles.subtitle}>
                  {moment(cert.Date).format("L")}
                </div>
              </div>
            </div>
          );
        })
      ) : (
        <div className={styles.nodatafound}>
          <Label>No certification found...</Label>
        </div>
      )}
    </div>
  );
};

export default Certification;
