import * as React from "react";
import { useMemo, useState, useEffect } from "react";
import { NewsService } from "../services/NewsService";
import ErrorService from "../services/ErrorService";


import {
  ButtonGroup,
  Card,
  Form,
  Image,
  Pagination,
  Table,
  ToggleButton,
  Button,
} from "react-bootstrap";
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
import styles from "./MyViewNews.module.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap-icons/font/bootstrap-icons.css";

import newsIcon from "../assets/news.png";
import eventsIcon from "../assets/events.png";
import broadcastIcon from "../assets/broadcasts.png";
import brandIcon from "../assets/brands.png";

import Container from "react-bootstrap/Container";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
import Tab from "react-bootstrap/Tab";
import Tabs from "react-bootstrap/Tabs";

const STATUSES = [
  "Draft",
  "Pending",
  "Published",
  "Approval Queue",
  "View Requests",
];

interface Props {
  context: any;
  items: any[];
  isLoading: boolean;
  selectedStatus: string;
  search: string;
  onStatusChange: (v: string) => void;
  onSearchChange: (v: string) => void;
  onEdit: (item: any) => void;
  onCreate: () => void;
  canCreateNews: boolean;
  onPreviewPopup: (item: any) => void;
  isSiteAdmin?: boolean;
}

const PAGE_SIZES = [10, 25, 50, 100];

const NewsTable: React.FC<Props> = ({
  context,
  items,
  isLoading,
  selectedStatus,
  search,
  onStatusChange,
  onSearchChange,
  onEdit,
  onCreate,
  canCreateNews,
  onPreviewPopup,
  isSiteAdmin,
}) => {
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(10);
  const [activeMainTab, setActiveMainTab] = useState("news");
  const [currentUserId, setCurrentUserId] = useState<number>(0);
  const [isUserLoading, setIsUserLoading] = useState(true);

  const newsService = React.useMemo(() => new NewsService(context), [context]);
  useEffect(() => {
    const loadCurrentUser = async () => {
      try {
        setIsUserLoading(true);
        const user = await newsService.sp.web.currentUser();
        setCurrentUserId(user.Id);
      } catch (error) {
ErrorService.logError(
  "NewsTable.loadCurrentUser",
  error
);  
      } finally {
        setIsUserLoading(false);
      }
    };

    loadCurrentUser();
  }, [newsService]);

  const [counts, setCounts] = useState({
    news: 0,
    events: 0,
    broadcasts: 0,
    brands: 0,
  });

  useEffect(() => {
    loadDashboardCounts();
  }, []);

  const loadDashboardCounts = async () => {
    const result = await newsService.getAllCounts();
    setCounts(result);
  };

  useEffect(() => {
    const url = window.location.pathname.toLowerCase();

    if (url.includes("events.aspx")) {
      setActiveMainTab("events");
    } else if (url.includes("broadcasts.aspx")) {
      setActiveMainTab("broadcasts");
    } else if (url.includes("brands.aspx")) {
      setActiveMainTab("brands");
    } else {
      setActiveMainTab("news");
    }
  }, []);

  /* ---------------- FILTER DATA ---------------- */
  const formatRequestNumber = (item: any) => {
    const id = item?.Id || item?.ID;
    if (!id) return "-";

    const padded = ("00000000" + id.toString()).slice(-8);
    return `RQ${padded}`;
  };

  const matchesStatusTab = (item: any, tab: string, userId: number) => {
    const s = (item.Status || "").trim().toLowerCase();

    /* ===================== DRAFT ===================== */
    // All items created by logged in user which are saved but not submitted
    if (tab === "Draft") {
      return s === "draft" && item.AuthorId === userId;
    }

    /* ===================== PENDING ===================== */
    // All items pending WITH logged in user for approval
    if (tab === "Pending") {
      return (
        item.PendingWithId === userId &&
        (s === "pendingwithlm" ||
          s === "pendingwithcm" ||
          s === "pendingwithcreator")
      );
    }

    /* ===================== APPROVAL QUEUE ===================== */
    // All items created by logged in user and still pending (with LM or CM)
    if (tab === "Approval Queue") {
      return (
        item.AuthorId === userId &&
        (s === "pendingwithlm" || s === "pendingwithcm")
      );
    }

    /* ===================== PUBLISHED ===================== */
    // Published items where user is creator OR part of workflow
    if (tab === "Published") {
      return (
        s === "published" &&
        (item.AuthorId === userId ||
          item.AssignToId === userId ||
          item.FollowOnActionById === userId ||
          item.PendingWithId === userId)
      );
    }

    /* ===================== VIEW REQUESTS ===================== */
    // All items where user is involved — independent of status
    if (tab === "View Requests") {
      return (
        (item.AuthorId === userId &&
          (s === "draft" ||
            s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "pendingwithcreator" ||
            s === "published" ||
            s === "rejected")) ||
        (item.AssignToId === userId &&
          (s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithcreator")) ||
        (item.FollowOnActionById === userId &&
          (s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithlm")) ||
        (item.PendingWithId === userId &&
          (s === "pendingwithlm" ||
            s === "pendingwithcm" ||
            s === "published" ||
            s === "rejected" ||
            s === "pendingwithcreator"))
      );
    }

    return false;
  };

  const isTableLoading = isLoading || isUserLoading;

  const filteredData = useMemo(() => {
    if (isTableLoading) return [];
    if (!items || !items.length) return [];

    return items.filter((item) => {
      /* ---------- STATUS FILTER ---------- */
      const matchesStatus = matchesStatusTab(
        item,
        selectedStatus,
        currentUserId,
      );

      /* ---------- SEARCH FILTER ---------- */
      const searchText = (search || "").trim().toLowerCase();

      const matchesSearch =
        !searchText ||
        (item.Title || "").toLowerCase().includes(searchText) ||
        formatRequestNumber(item).toLowerCase().includes(searchText);

      return matchesStatus && matchesSearch;
    });
  }, [items, selectedStatus, search, currentUserId, isTableLoading]);

  const getTabStatusName = (itemStatus: string) => {
    const s = (itemStatus || "").trim();

    switch (s.toLowerCase()) {
      case "draft":
        return "Draft";

      case "pendingwithcreator":
        return "PendingWithCreator";

      case "pendingwithlm":
        return "PendingWithLM";

      case "pendingwithcm":
        return "PendingWithCM";

      case "published":
        return "Published";

      case "rejected":
        return "Rejected";

      case "deleted":
        return "Deleted";

      case "closed":
        return "Closed";

      case "canceled":
        return "Canceled";

      default:
        return s; // fallback show original
    }
  };

  /* ---------------- PAGINATION ---------------- */
  const totalItems = filteredData.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const startIndex = (currentPage - 1) * pageSize;

  const pageItems = filteredData.slice(startIndex, startIndex + pageSize);

  //  Fix page number if pageSize change makes currentPage invalid
  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages === 0 ? 1 : totalPages);
    }
  }, [pageSize, totalPages]);

  const formatISTDate = (dateValue: string) => {
    if (!dateValue) return "-";

    return new Date(dateValue).toLocaleDateString("en-GB", {
      timeZone: "Asia/Kolkata",
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  };

  return (
    <div> 
      <Container fluid className={styles.dashboardWrapper}>
        {/* Image + overlay section */}
        <div className={styles.topImageSection}>
          <div className={styles.topImageContent}>
            {/* DASHBOARD CARDS */}
            <Row className="my-2">
              <div className="d-flex justify-content-between">
                <h5 className={styles.pageHeadingWhite}>CPH Dashboard</h5>
                <div>
                  {isSiteAdmin && (
                    <Button
                      variant="outline-primary"
                      className={styles.navigateBtn}
                      onClick={() =>
                        window.open(
                          `${context.pageContext.web.absoluteUrl}/SitePages/CPHAdminDashboard.aspx`,
                          "_blank",
                        )
                      }
                    >
                      <span className="ml-2">Admin </span>
                      <i className="bi bi-arrow-right" />
                    </Button>
                  )}
                </div>
              </div>
            </Row>
            <Row className="g-3 justify-content-between">
              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={`${styles.cardCommon}`}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total News
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.news}
                      </Card.Title>
                    </div>
                    <img
                      alt="News"
                      src={newsIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={`${styles.cardCommon}`}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Events
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.events}
                      </Card.Title>
                    </div>
                    <img
                      alt="Events"
                      src={eventsIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={`${styles.cardCommon}`}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Broadcasts
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.broadcasts}
                      </Card.Title>
                    </div>
                    <img
                      alt="Broadcasts"
                      src={broadcastIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              <Col xs={12} sm={6} md={4} lg={4} xl={3} xxl={3}>
                <Card className={`${styles.cardCommon}`}>
                  <Card.Body className="d-flex justify-content-between">
                    <div>
                      <Card.Subtitle className={styles.dashboardCardSubTitle}>
                        Total Brands
                      </Card.Subtitle>
                      <Card.Title className={styles.dashboardCardCount}>
                        {counts.brands}
                      </Card.Title>
                    </div>
                    <img
                      alt="Brands"
                      src={brandIcon}
                      className={styles.dashboardCardIcon}
                    />
                  </Card.Body>
                </Card>
              </Col>

              {/* Cards omitted for brevity — unchanged */}
            </Row>
          </div>
        </div>
        <Row className="mt-4 mx-0 mx-md-1">
          <Col className="px-0 px-md-1">
            <Card className={`${styles.cardCommon} p-0 p-md-2`}>
              {/* ---------------- BODY ---------------- */}
              <div className="card-body">
                <Tabs
                  activeKey={activeMainTab}
                  onSelect={(k) => {
                    if (!k) return;

                    if (k === "news") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/CPH_Dashboard.aspx`;
                      return;
                    }

                    if (k === "events") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Events.aspx`;
                      return;
                    }

                    if (k === "broadcasts") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Broadcast.aspx`;
                      return;
                    }

                    if (k === "brands") {
                      window.location.href = `${context.pageContext.web.absoluteUrl}/SitePages/Brand.aspx`;
                      return;
                    }
                  }}
                  justify
                  className={styles.customTabs}
                >
                  <Tab eventKey="news" title="News">
                    {/* <div className="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-2"> */}
                    <div className="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-4 mt-4 w-100">
                      {/* LEFT — Status buttons */}
                      <ButtonGroup
                        className={`${styles.statusGroup} flex-wrap`}
                      >
                        {STATUSES.map((status) => (
                          <ToggleButton
                            key={status}
                            id={`status-${status}`}
                            type="radio"
                            variant="outline-primary"
                            name="statusFilter"
                            value={status}
                            checked={selectedStatus === status}
                            onChange={(e) => {
                              onStatusChange(e.currentTarget.value);
                              setCurrentPage(1);
                            }}
                            className={styles.statusBtn}
                          >
                            {status}
                          </ToggleButton>
                        ))}
                      </ButtonGroup>

                      {/* RIGHT — Search + Icons + Create */}
                      <div
                        className="d-flex align-items-center gap-2 ms-auto mt-1 mt-md-0"
                        style={{ height: 44 }}
                      >
                        <div className={`input-group ${styles.searchBox}`}>
                          <span className="input-group-text bg-white border-end-0">
                            <i className="bi bi-search" />
                          </span>

                          <Form.Control
                            type="text"
                            placeholder="Search..."
                            value={search}
                            onChange={(e) => {
                              onSearchChange(e.target.value);
                              setCurrentPage(1);
                            }}
                            className={`border-start-0 ${styles.noFocus}`}
                          />
                        </div>

                        {canCreateNews && (
                          <Button
                            className={`mt-3 ${styles.createNewsBtn}`}
                            onClick={onCreate}
                          >
                            <i className="bi bi-plus-circle me-1" />
                            Create News
                          </Button>
                        )}
                      </div>
                    </div>

                    {/* TABLE */}
                    <Table responsive className={styles.table}>
                      <thead>
                        <tr>
                          <th>Request Number</th>
                          <th>Image</th>
                          <th>Title</th>
                          <th>Event Date</th>
                          <th>Created On</th>
                          <th>Created By</th>
                          <th>Pending With</th>
                          <th>Status</th>
                          <th>Action</th>
                        </tr>
                      </thead>

                      {/* <tbody>
                        {pageItems.length === 0 && (
                          <tr>
                            <td colSpan={8} className="text-center py-4">
                              No records found
                            </td>
                          </tr>
                        )}

                        {pageItems.map((item, index) => (
                          <tr key={item.Id || index}>
                            <td>
                              <a
                                href="#"
                                className={styles.documentLink}
                                onClick={(e) => {
                                  e.preventDefault();
                                  onPreviewPopup(item);
                                }}
                              >
                                {formatRequestNumber(item)}
                              </a>
                            </td>

                            <td>
                              {item.ThumbnailUrl ? (
                                <Image
                                  src={item.ThumbnailUrl}
                                  rounded
                                  width={50}
                                  height={50}
                                  style={{ objectFit: "cover" }}
                                />
                              ) : (
                                <span>-</span>
                              )}
                            </td>

                            <td className={styles.titleCol} title={item.Title}>
                              {item.Title}
                            </td>

                           
                            <td>{formatISTDate(item.EventDate)}</td>
<td>{formatISTDate(item.Created)}</td>

                            <td>{item.Author?.Title}</td>
                            <td>{item.PendingWith?.Title}</td>
                            <td>{getTabStatusName(item.Status)}</td>

                            <td className={styles.tableIcon}>
                              <i
                                className="bi bi-pencil-square"
                                onClick={() => onEdit(item)}
                                style={{ cursor: "pointer" }}
                              />
                            </td> 
                          </tr>
                        ))}
                      </tbody> */}
                      <tbody>
                        {isTableLoading ? (
                          <tr>
                            <td colSpan={9} className="text-center py-4">
                              <div className="d-flex justify-content-center align-items-center gap-2">
                                <div
                                  className="spinner-border spinner-border-sm text-primary"
                                  role="status"
                                />
                                <span>Loading data...</span>
                              </div>
                            </td>
                          </tr>
                        ) : pageItems.length === 0 ? (
                          <tr>
                            <td colSpan={9} className="text-center py-4">
                              No records found
                            </td>
                          </tr>
                        ) : (
                          pageItems.map((item, index) => (
                            <tr key={item.Id || index}>
                              <td>
                                <a
                                  href="#"
                                  className={styles.documentLink}
                                  onClick={(e) => {
                                    e.preventDefault();
                                    onPreviewPopup(item);
                                  }}
                                >
                                  {formatRequestNumber(item)}
                                </a>
                              </td>

                              <td>
                                {item.ThumbnailUrl ? (
                                  <Image
                                    src={item.ThumbnailUrl}
                                    rounded
                                    width={50}
                                    height={50}
                                    style={{ objectFit: "cover" }}
                                  />
                                ) : (
                                  <span>-</span>
                                )}
                              </td>

                              <td
                                className={styles.titleCol}
                                title={item.Title}
                              >
                                {item.Title}
                              </td>

                              <td>{formatISTDate(item.EventDate)}</td>
                              <td>{formatISTDate(item.Created)}</td>
                              <td>{item.Author?.Title}</td>
                              <td>{item.PendingWith?.Title}</td>
                              <td>{getTabStatusName(item.Status)}</td>

                              <td className={styles.tableIcon}>
                                <i
                                  className="bi bi-pencil-square"
                                  onClick={() => onEdit(item)}
                                  style={{ cursor: "pointer" }}
                                />
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </Table>

                    {/* FOOTER */}
                    <div className={styles.tableFooter}>
                      <div className={styles.entriesInfo}>
                        Showing {totalItems === 0 ? 0 : startIndex + 1}–
                        {Math.min(startIndex + pageSize, totalItems)} of{" "}
                        {totalItems} entries
                      </div>

                      {/*  RIGHT SIDE GROUP */}
                      <div className="d-flex align-items-center gap-3">
                        {/* Entries per page */}
                        <div className="d-flex align-items-center gap-2">
                          <Form.Select
                            style={{ width: "90px" }}
                            value={pageSize}
                            onChange={(e) => {
                              setPageSize(Number(e.target.value));
                              setCurrentPage(1);
                            }}
                          >
                            {PAGE_SIZES.map((n) => (
                              <option key={n} value={n}>
                                {n}
                              </option>
                            ))}
                          </Form.Select>
                          <span className="mb-0">entries per page</span>
                        </div>

                        {/* Pagination */}
                        <Pagination
                          size="sm"
                          className={`${styles.customPagination} d-flex align-items-center mb-0`}
                        >
                          <Pagination.Prev
                            disabled={currentPage === 1}
                            onClick={() => setCurrentPage((p) => p - 1)}
                          />

                          {[...Array(totalPages)].map((_, i) => (
                            <Pagination.Item
                              key={i}
                              active={currentPage === i + 1}
                              onClick={() => setCurrentPage(i + 1)}
                            >
                              {i + 1}
                            </Pagination.Item>
                          ))}

                          <Pagination.Next
                            disabled={
                              currentPage === totalPages || totalPages === 0
                            }
                            onClick={() => setCurrentPage((p) => p + 1)}
                          />
                        </Pagination>
                      </div>
                    </div>
                  </Tab>
                  <Tab eventKey="events" title="Events">
                    Tab content for Events
                  </Tab>
                  <Tab eventKey="broadcasts" title="Broadcasts">
                    Tab content for broadcasts
                  </Tab>
                  <Tab eventKey="brands" title="Brands">
                    Tab content for brands
                  </Tab>
                </Tabs>
              </div>
            </Card>
          </Col>
        </Row>
      </Container>
    </div>
    // </div>
  );
};

export default NewsTable;
