import * as React from "react";
import styles from "./CertDashboard.module.scss";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { NormalPeoplePicker } from "@fluentui/react";
import { InputText } from "primereact/inputtext";

function CertDashboard(props: any) {
  const [certifications, setCertifications] = useState([]);
  const [allcertifications, setAllCertifications] = useState([]);
  const rows = 10;
  const [first, setFirst] = useState(0);
  const [sortFields, setSortField] = useState<string>("");
  const [sortOrder, setSortOrder] = useState(null);
  const [count, setCount] = useState(0);

  const [filterDataAll, setFilterData] = useState<any>({
    searchValue: "",
    dropValue: null,
  });

  console.log(allcertifications, "all");
  const onPageChange = (event: any) => {
    setFirst(event.first);
    certifications.slice(first, first + rows);
  };
  const onSort = (e: any) => {
    setSortField(e.sortField);
    setSortOrder(e.sortOrder);
  };

  const dateformate = (rowData: any) => {
    // Assuming your date field in the row is 'Date'
    return rowData.Date ? moment(rowData.Date).format("DD/MM/YYYY") : "";
  };
  const findUser = (user: any) => {
    let finduser: any = [];
    if (user) {
      finduser = [...props.SiteUsers].filter((val) => {
        return val.ID == user;
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

  // const handleSearch = (name: any, value: any) => {

  //   let filterData = [...allcertifications];

  //   if (name == "Search") {
  //     filterData = filterData.filter((item: any) => {
  //       const searchValue = value.toLowerCase().replace(/\s+/g, "");

  //       return (
  //         item.CertificationName?.toLowerCase()
  //           .replace(/\s+/g, "")
  //           .includes(searchValue) ||
  //         dateformate(item).replace(/\s+/g, "").includes(searchValue) ||
  //         item.Department?.toLowerCase()
  //           .replace(/\s+/g, "")
  //           .includes(searchValue) ||
  //         item.EmployeeName?.toLowerCase()
  //           .replace(/\s+/g, "")
  //           .includes(searchValue)
  //       );
  //     });

  //     setFilterData({ ...filterDataAll, searchValue: value });
  //   }if(name == "Drop") {
  //     if (value) {
  //       filterData = filterData.filter((item: any) => {
  //         return item.EmployeeNameID == value;
  //       });
  //     }

  //     setFilterData({ ...filterDataAll, dropValue: value });
  //   }

  //   setCertifications(filterData);
  // };

  const handleSearch = (name: any, value: any) => {
    let updatedFilters = { ...filterDataAll };

    if (name === "Search") {
      updatedFilters.searchValue = value;
    } else if (name === "Drop") {
      updatedFilters.dropValue = value;
    }

    // Apply both filters together
    let filterData = [...allcertifications];

    // 1. Search filter
    if (updatedFilters.searchValue) {
      const searchValue = updatedFilters.searchValue
        .toLowerCase()
        .replace(/\s+/g, "");
      filterData = filterData.filter((item: any) => {
        return (
          item.CertificationName?.toLowerCase()
            .replace(/\s+/g, "")
            .includes(searchValue) ||
          dateformate(item).replace(/\s+/g, "").includes(searchValue) ||
          item.Department?.toLowerCase()
            .replace(/\s+/g, "")
            .includes(searchValue) ||
          item.EmployeeName?.toLowerCase()
            .replace(/\s+/g, "")
            .includes(searchValue)
        );
      });
    }

    // 2. Dropdown filter
    if (updatedFilters.dropValue) {
      filterData = filterData.filter(
        (item: any) => item.EmployeeNameID == updatedFilters.dropValue
      );
    }

    setFilterData(updatedFilters);
    setCertifications(filterData);
  };
  const SerachClear = () => {
    setFilterData({ searchValue: "", dropValue: null });

    setCertifications(allcertifications);
    setCount((count) => count + 1);
    console.log(filterDataAll, "filterDataAll");
  };

  useEffect(() => {
    setCertifications(props.Certifications);
    setAllCertifications(props.Certifications);
  }, []);

  return (
    <>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "10px",
          marginBottom: "20px",
          justifyContent: "space-between",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div
            onClick={() => {
              props.setPageConfig({ page: "Dashboard" });
            }}
          >
            <i className="pi pi-arrow-left" style={{ marginTop: "4px" }}></i>
          </div>

          <h4 className={styles.heading}>Certifications</h4>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
       
          <NormalPeoplePicker
            onResolveSuggestions={GetUserDetails}
            inputProps={{
              placeholder: "Select a employee",
            }}
            itemLimit={1}
            defaultSelectedItems={findUser(filterDataAll.dropValue)}
            onChange={(SelectUser: any) => {
              {
                SelectUser.length
                  ? handleSearch("Drop", SelectUser[0]["ID"])
                  : handleSearch("Drop", null);
              }
            }}
            key={count}
          />

          <div>
            <InputText
              className={styles.search}
              placeholder="Search"
              value={filterDataAll.searchValue}
              onChange={(event: any) => {
                handleSearch("Search", event.target.value);
              }}
            />
          </div>

          <i
            style={{
             
              background: "#3f9f4e",
              color:"white",
              borderRadius:"7px",
              textAlign:"center",
              paddingTop:"9px",
              width: "40px",
              height: "36px",
              cursor: "pointer",
            }}
            className="pi pi-sync"
            onClick={() => {
              void SerachClear();
            }}
          ></i>
        </div>
      </div>

      <div className={styles.Dashboard_container}>
        <div
          className="data-table-template deal-tracker-data-table"
          style={{ marginTop: "8px" }}
        >
          <DataTable
            value={certifications}
            rows={rows}
            paginator={certifications.length > rows}
            totalRecords={certifications.length}
            onPage={(e) => onPageChange(e)}
            first={first}
            paginatorRight={null}
            paginatorLeft={`Showing ${
              certifications.length === 0 ? 0 : first + 1
            } to ${Math.min(first + rows, certifications.length)} of ${
              certifications.length
            } records`}
            emptyMessage={
              <>
                <p style={{ textAlign: "center" }}>List is empty</p>
              </>
            }
            removableSort
            sortOrder={sortOrder}
            sortField={sortFields}
            onSort={onSort}
          >
            <Column
              field="EmployeeName"
              sortable
              header="Employee Name"
              style={{ width: "15%" }}
            />
            <Column
              field="CertificationName"
              sortable
              header="Certification Name"
              style={{ width: "20%" }}
            />
            <Column
              field="Department"
              sortable
              header="Department"
              style={{ width: "15%" }}
            />
            <Column
              field="Date"
              body={(rowData) =>
                rowData.Date ? moment(rowData.Date).format("DD/MM/YYYY") : ""
              }
              sortable
              header="Date"
              style={{ width: "15%" }}
            />
          </DataTable>
        </div>
      </div>
    </>
  );
}
export default CertDashboard;
