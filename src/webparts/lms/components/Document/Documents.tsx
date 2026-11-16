import { Label } from "@fluentui/react";
import styles from "./Document.module.scss";
import * as React from "react";
import { Tooltip } from "primereact/tooltip";

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

function Documents(props: any) {
  const siteUrl = props.context.pageContext.web.serverRelativeUrl;
  const baseLibraryPath = `${siteUrl}/Shared Documents`;
  const [allDocuments, setAllDocuments] = React.useState<any[]>([]);
  const [currentFolderUrl, setCurrentFolderUrl] =
    React.useState<string>(baseLibraryPath);
  const [folderStack, setFolderStack] = React.useState<string[]>([]);
  function _sortFilterKeys(a: any, b: any) {
    if (a.time > b.time) {
      return -1;
    }
    if (a.time < b.time) {
      return 1;
    }
    return 0;
  }
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

    setAllDocuments(docs.sort(_sortFilterKeys).slice(0, 7));
  }, [props.Document]);
  const filteredFolders = props.Document.filter(
  (item: any) =>
    item.FolderUrl.startsWith(currentFolderUrl + "/") &&
    item.supportingDocuments &&
    item.supportingDocuments.length > 0
);

// Dashboard -> 7 folders
// Documents page -> All folders
const displayedFolders = props.isvisible
  ? filteredFolders
  : filteredFolders.slice(0, 7);

  const currentFolderFiles = props.Document.filter(
    (f: any) => f.FolderUrl === currentFolderUrl
  ).flatMap((f: any) => f.supportingDocuments);

  const handleFolderClick = (folder: any) => {
    setFolderStack([...folderStack, currentFolderUrl]);
    setCurrentFolderUrl(folder.FolderUrl);
  };
  const normalizeUrl = (url: string) => {
    if (!url) return "";
    return decodeURIComponent(url.trim().toLowerCase().replace(/\/+$/, "")); // removes trailing slash
  };
  const breadcrumbParts = React.useMemo(() => {
    const normalizedBase = normalizeUrl(baseLibraryPath);
    const normalizedCurrent = normalizeUrl(currentFolderUrl);
    const relativePath = normalizedCurrent.replace(normalizedBase, "");
    const parts = relativePath.split("/").filter(Boolean);
    return parts;
  }, [currentFolderUrl, baseLibraryPath]);

  return (
    <>
      <div className={styles.recentDocuments}>
        <div className={styles.header}>
          <h4 className={styles.title}>Recent Document</h4>

          {allDocuments.length > 0 && (
            <a
              className={styles.seeAll}
              onClick={() => {
                props.setPageConfig({ page: "Documents" });
                props.setIsvisible(false);
              }}
            >
              See all
            </a>
          )}
        </div>

        <div className={styles.allDocuments}>
          <div className={styles.breadcrumb}>
            <span
              style={{
                color: "#000",
                cursor: breadcrumbParts.length === 0 ? "default" : "pointer",
                fontWeight: 400,
              }}
              onClick={() => setCurrentFolderUrl(baseLibraryPath)}
            >
              All Documents
            </span>

            {breadcrumbParts.length > 0 && (
              <>
                <span style={{ margin: "0 6px", color: "#666" }}>›</span>
                <>
                  <Tooltip target=".breadcrumb-folder-name" position="top" />
                  <span
                    className="breadcrumb-folder-name"
                    data-pr-tooltip={decodeURIComponent(
                      breadcrumbParts[breadcrumbParts.length - 1]
                    )}
                    style={{
                      color: "#0078D4",
                      fontWeight: 600,
                      width: "250px",
                      overflow: "hidden",
                      display: "-webkit-box",
                      WebkitLineClamp: 2,
                      WebkitBoxOrient: "vertical",
                      textOverflow: "ellipsis",
                      whiteSpace: "normal",
                      cursor: "default",
                    }}
                  >
                    {decodeURIComponent(
                      breadcrumbParts[breadcrumbParts.length - 1]
                    )}
                  </span>
                </>
              </>
            )}
          </div>

          {/* Folder & File List */}
          <div className={styles.folderList}>
            {displayedFolders.map((folder: any, idx: number) => (
              <div
                key={folder.FolderUrl}
                className={styles.folderItem}
                onClick={() => handleFolderClick(folder)}
              >
                <Tooltip target={`.folder-name-${idx}`} />
                <img
                  src={logos.another}
                  alt="folder"
                  className={styles.fileIcon}
                />
                <span
                  data-pr-tooltip={folder.FolderName}
                  data-pr-position="bottom"
                  className={`folder-name-${idx} ${styles.documentName}`}
                >
                  {folder.FolderName}
                </span>
              </div>
            ))}

            {currentFolderFiles.length > 0 &&
              currentFolderFiles.map((file: any, idx: number) => (
                <div key={idx} className={styles.fileItem}>
                  <img
                    src={getImageUrl(file.name)}
                    alt="file"
                    className={styles.fileIcon}
                  />
                  <Tooltip target={`.file-name-${idx}`} />
                  <a
                    href={file.url}
                    className={`file-name-${idx} ${styles.documentName}`}
                    target="_blank"
                    rel="noopener noreferrer"
                    data-pr-tooltip={file.name}
                    data-pr-position="top"
                  >
                    {file.name}
                  </a>
                </div>
              ))}

            {filteredFolders.length === 0 &&
              currentFolderFiles.length === 0 && (
                <div className={styles.nodatafound}>
                  <Label>No items found in this folder</Label>
                </div>
              )}
          </div>
        </div>
      </div>
    </>
  );
}

export default Documents;
