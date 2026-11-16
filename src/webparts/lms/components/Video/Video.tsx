import * as React from "react";
import styles from "./Video.module.scss"; // CSS Module import
import * as moment from "moment";
import { Label } from "@fluentui/react";
import { Tooltip } from "primereact/tooltip";

interface IVideo {
  name: string;
  url: string;
  type: string;
  TimeCreated: string;
}

const VideoComponent = (props: any) => {
  const [videoDetails, setVideoDetails] = React.useState<IVideo[]>([]);

  const [selectedVideo, setSelectedVideo] = React.useState<IVideo | null>(null);
  const [showPopup, setShowPopup] = React.useState(false);

  React.useEffect(() => {
    setVideoDetails(props.videoDetails.slice(0, 5));
  }, [props.videoDetails]);

  return (
    <>
      <div style={{ display: "flex" }}>
        <div className={styles.recentVideos}>
          <div className={styles.header}>
            <h4 className={styles.title}>Recent Videos</h4>

            {videoDetails.length > 0 && (
              <a
                className={styles.seeAll}
                onClick={() => {
                  props.setPageConfig({ page: "Documents" });
                  props.setIsvisible(true);
                }}
              >
                See all
              </a>
            )}
          </div>
          <div className={styles.videoContainer}>
            {videoDetails.map((video, index) => (
              <div className={styles.videoCard} key={index}>
                <div
                  style={{ display: "flex", alignItems: "center", gap: "15px" }}
                >
                  {/* <img src={window.location.origin + video.url}></img> */}
                  <div
                    className="video-thumb"
                    onClick={() => {
                      setSelectedVideo(video);
                      setShowPopup(true);
                    }}
                  >
                    <video muted>
                      <source src={video.url} />
                    </video>
                    <div className="play-overlay">▶</div>
                  </div>
                  <div className={styles.textContainer}>
                    <Tooltip target={`.file-name-${index}`} />
                    <a
                      href=""
                      onClick={() => {
                        setSelectedVideo(video);
                        setShowPopup(true);
                      }}
                      rel="noopener noreferrer"
                      className={`file-name-${index} ${styles.documentName}`}
                      data-pr-tooltip={video.name.replace(/\.[^/.]+$/, "")}
                       data-pr-position="top"
                    >
                      {video.name?.replace(/\.[^/.]+$/, "")}
                    </a>
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
                      {moment(video.TimeCreated).fromNow()}
                    </div>
                  </div>
                </div>

                <div className={styles.videoContaineDownload}>
                  <a
                    href={video.url}
                    download={video.name}
                    className={styles.downloadBtn}
                    title=""
                  >
                    <i className="pi pi-download" />
                  </a>
                </div>
              </div>
            ))}
            {videoDetails.length == 0 && (
              <div className={styles.nodatafound}>
                <Label
                  style={{
                    padding: "5px",
                    paddingLeft: "20px",
                    textAlign: "center",
                  }}
                >
                  No video found...
                </Label>
              </div>
            )}

            {showPopup && selectedVideo && (
              <div className={styles.modalOverlay}>
                <div className={styles.modalContent}>
                  <h3>{selectedVideo.name}</h3>
                  <iframe
                    src={selectedVideo.url}
                    className={styles.iframePlayer}
                    allow="autoplay"
                    allowFullScreen
                  ></iframe>

                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setSelectedVideo(null);
                      setShowPopup(false);
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
        <div
          className="vertical-line-container"
          style={{ lineHeight: "200px" }}
        >
          <div className="line"></div>
        </div>
      </div>
    </>
  );
};

export default VideoComponent;
