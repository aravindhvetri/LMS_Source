import * as React from "react";
import "../../../../ExternalRef/CSS/Style.css";
import styles from "./DocumentDashBoard.module.scss";
import { Tooltip } from "primereact/tooltip";
import { Label } from "@fluentui/react/lib/Label";
import * as moment from "moment";

interface FileItem {
  Types?: string;
  url: string;
  name: string;
  time?: string;
  CreatedByName?: string;
}

interface DocumentFolder {
  FolderName: string;
  FolderUrl: string;
  time?: string;
  CreatedByName?: string;
  supportingDocuments: FileItem[];
}

interface IVideo {
  name: string;
  url: string;
  type: string;
  TimeCreated: string;
}

function Documents(props: any) {
  const logos = {
    pdf: require("../../components/assets/pdf.png"),
    ppt: require("../../components/assets/ppt.png"),
    excel: require("../../components/assets/xls.png"),
    word: require("../../components/assets/doc.png"),
    image: require("../../components/assets/image.png"),
    another: require("../../components/assets/folder.png"),
  };

  const getImageUrl = (filename: string) => {
    const ext = filename.split(".").pop()?.toLowerCase();
    if (!ext) return logos.another;
    if (ext === "pdf") return logos.pdf;
    if (["ppt", "pptx", "ppsx"].includes(ext)) return logos.ppt;
    if (["xls", "xlsx", "csv"].includes(ext)) return logos.excel;
    if (["doc", "docx"].includes(ext)) return logos.word;
    if (["png", "svg", "jpg", "jpeg"].includes(ext)) return logos.image;
    return logos.another;
  };

  const [documents, setDocuments] = React.useState<DocumentFolder[]>([]);
  const [video, setVideo] = React.useState<any[]>([]);
  const [selectedVideo, setSelectedVideo] = React.useState<IVideo | null>(null);
  const [showPopup, setShowPopup] = React.useState(false);

  const siteUrl = props.context.pageContext.web.serverRelativeUrl;
  const baseLibraryPath = `${siteUrl}/Shared Documents`;
  const [currentFolderUrl, setCurrentFolderUrl] =
    React.useState<string>(baseLibraryPath);
  const [folderStack, setFolderStack] = React.useState<string[]>([]);

  const downloadFiles = (file: any) => {
    const link = document.createElement("a");
    link.href = file.url.startsWith("http")
      ? file.url
      : window.location.origin + file.url;
    link.download = file.name || "download";
    link.target = "_blank";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const filteredFolders = props.Document.filter(
    (item: any) =>
      item.FolderUrl.startsWith(currentFolderUrl + "/") &&
      item.FolderUrl.replace(currentFolderUrl + "/", "").split("/").length ===
        1,
  );

  const currentFolderFiles = props.Document.filter(
    (f: any) => f.FolderUrl === currentFolderUrl,
  ).flatMap((f: any) => f.supportingDocuments);

  const handleFolderClick = (folder: any) => {
    setFolderStack([...folderStack, currentFolderUrl]);
    setCurrentFolderUrl(folder.FolderUrl);
  };

  const breadcrumbParts = currentFolderUrl
    .replace(baseLibraryPath, "")
    .split("/")
    .filter(Boolean);

  const handleBreadcrumbClick = (index: number) => {
    const newPath = `${baseLibraryPath}/${breadcrumbParts
      .slice(0, index + 1)
      .join("/")}`;
    setCurrentFolderUrl(newPath);
  };
  console.log(documents);
  React.useEffect(() => {
    const docs: any[] = [];
    props.Document.forEach((files: any) => {
      files.supportingDocuments?.forEach((file: any) => {
        docs.push({
          name: file.name,
          url: file.url,
          type: file.type,
          time: file.time,
          CreatedByName: file.CreatedByName,
        });
      });
    });
    setVideo(props.videoDetails);
    setDocuments(props.Document);
  }, [props.Document]);

  return (
    <>
      {/* Header */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "10px",
          marginBottom: "20px",
        }}
      >
        <div
          onClick={() => {
            props.setPageConfig({ page: "Dashboard" });
          }}
        >
          <i className="pi pi-arrow-left" style={{ marginTop: "4px" }}></i>
        </div>
        <h4 className={styles.heading}>
          Events {props.isVideo ? "Videos" : "Documents"}
        </h4>
      </div>

      {/* Video Section */}
      {props.isVideo ? (
        <div className={styles.videoContainer}>
          {video.map((video, index) => (
            <div className={styles.videoCard} key={index}>
              <div
                style={{ display: "flex", alignItems: "center", gap: "15px" }}
              >
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
                    href="#"
                    onClick={() => {
                      setSelectedVideo(video);
                      setShowPopup(true);
                    }}
                    rel="noopener noreferrer"
                    className={`file-name-${index} ${styles.documentName}`}
                    data-pr-tooltip={video.name}
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
                  title="Download"
                >
                  <i className="pi pi-download" />
                </a>
              </div>
            </div>
          ))}

          {video.length === 0 && (
            <div className={styles.nodatafound}>
              <Label>No video found...</Label>
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
      ) : (
        <>
          {/* Document Section */}
          <div>
            <div className={styles.breadcrumb}>
              <span
                style={{
                  color: "black",
                  cursor: "pointer",
                  fontWeight: 500,
                  textDecoration: "underline",
                }}
                onClick={() => setCurrentFolderUrl(baseLibraryPath)}
              >
                All Documents
              </span>

              {breadcrumbParts.map((part, index) => (
                <span key={index}>
                  <span style={{ margin: "0 6px", color: "#666" }}>›</span>
                  <span
                    style={{
                      cursor:
                        index < breadcrumbParts.length - 1
                          ? "pointer"
                          : "default",
                      color:
                        index === breadcrumbParts.length - 1
                          ? "#0078D4"
                          : "#000",
                      fontWeight:
                        index === breadcrumbParts.length - 1 ? 600 : 400,
                    }}
                    onClick={() =>
                      index < breadcrumbParts.length - 1 &&
                      handleBreadcrumbClick(index)
                    }
                  >
                    {part}
                  </span>
                </span>
              ))}

              {/* Table Header */}
              <div className={styles.folderList}>
                <div className={styles.documentTable}>
                  <div className={styles.tableHeader}>
                    <div className={styles.colName}>Name</div>
                    <div className={styles.colModified}>Modified</div>
                    <div className={styles.colModifiedBy}>Modified by</div>
                  </div>

                  {/* 🔹 Folder Rows */}
                  {filteredFolders.map((folder: any, idx: number) => (
                    <div
                      key={folder.FolderUrl}
                      className={styles.tableRow}
                      onClick={() => handleFolderClick(folder)}
                      style={{ cursor: "pointer" }}
                    >
                      <div className={styles.colName}>
                        <img
                          src={logos.another}
                          alt="folder"
                          className={styles.fileIcon}
                        />
                        <span
                          data-pr-tooltip={folder.FolderName}
                          data-pr-position="bottom"
                          className={`folder-name-${idx} ${styles.documentName}`}
                          style={{ color: "black", cursor: "pointer" }}
                        >
                          {folder.FolderName}
                        </span>
                      </div>

                      <div className={styles.colModified}>
                        {folder.time
                          ? new Date(folder.time).toLocaleDateString("en-US", {
                              month: "short",
                              day: "numeric",
                            })
                          : "-"}
                      </div>

                      <div className={styles.colModifiedBy}>
                        {folder.CreatedByName || "-"}
                      </div>
                    </div>
                  ))}
                  {currentFolderFiles.length > 0 &&
                    currentFolderFiles
                      .slice(0, 7)
                      .map((file: any, idx: number) => (
                        <div key={idx} className={styles.tableRow}>
                          <div className={styles.colName}>
                            <img
                              src={getImageUrl(file.name)}
                              alt="file"
                              className={styles.fileIcon}
                            />
                            <span
                              onClick={() => downloadFiles(file)}
                              style={{ cursor: "pointer", color: "black" }}
                            >
                              {file.name}
                            </span>
                          </div>

                          <div className={styles.colModified}>
                            {file.time
                              ? new Date(file.time).toLocaleDateString(
                                  "en-US",
                                  {
                                    month: "short",
                                    day: "numeric",
                                  },
                                )
                              : "-"}
                          </div>

                          <div className={styles.colModifiedBy}>
                            {file.CreatedByName || "-"}
                          </div>
                        </div>
                      ))}

                  {/* No Data Message */}
                  {filteredFolders.length === 0 &&
                    currentFolderFiles.length === 0 && (
                      <div className={styles.nodatafound}>
                        <Label>No items found in this folder</Label>
                      </div>
                    )}
                </div>
              </div>
            </div>
          </div>
        </>
      )}
    </>
  );
}

export default Documents;
