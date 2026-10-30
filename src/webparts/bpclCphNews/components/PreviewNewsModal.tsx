import * as React from "react";
import { Carousel, Modal } from "react-bootstrap";
import styles from "./MyViewNews.module.scss";
//import "@fontsource/inter"; // default (400)
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";

interface Props {
  show: boolean;
  onClose: () => void;
  item: any;
  attachments?: { FileName: string; ServerRelativeUrl: string }[];
} 

const PreviewNewsModal: React.FC<Props> = ({
  show,
  onClose,
  item,
  attachments = [],
}) => {
  if (!item) return null;

  //  Build image list from attachments (Thumbnail + Picture1/2/3)
  const getFileUrl = (fileName: string) => {
    const f = attachments.find((x) => x.FileName === fileName);
    return f?.ServerRelativeUrl || "";
  };

  // const images: string[] = [];

  // if (item.Thumbnail) {
  //   const url = getFileUrl(item.Thumbnail);
  //   if (url) images.push(url + `?v=${new Date().getTime()}`);
  // }
  // if (item.Picture1) images.push(getFileUrl(item.Picture1));
  // if (item.Picture2) images.push(getFileUrl(item.Picture2));
  // if (item.Picture3) images.push(getFileUrl(item.Picture3));

  // const filteredImages = images.filter((x) => x);
const images: { url: string; caption: string }[] = [];
 
if (item.Thumbnail) {
  const url = getFileUrl(item.Thumbnail);
  if (url) {
    images.push({
      url: url + `?v=${new Date().getTime()}`,
      caption: item.ThumbnailCaption || "",
    });
  }
}

if (item.Picture1) {
  const url = getFileUrl(item.Picture1);
  if (url) {
    images.push({
      url,
      caption: item.Pic1Caption || "",
    });
  }
}

if (item.Picture2) {
  const url = getFileUrl(item.Picture2);
  if (url) {
    images.push({
      url,
      caption: item.Pic2Caption || "",
    });
  }
}

if (item.Picture3) {
  const url = getFileUrl(item.Picture3);
  if (url) {
    images.push({
      url,
      caption: item.Pic3Caption || "",
    });
  }
}

const filteredImages = images.filter((x) => x.url);

  const formatISTDate = (dateValue: string) => {
  if (!dateValue) return "-";

  return new Date(dateValue).toLocaleDateString("en-GB", {
    timeZone: "Asia/Kolkata", 
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};
 
const normalizeQuillHtml = (html: string): string => {
  if (!html) return "";

  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = html;

  // Remove empty paragraphs
  tempDiv.querySelectorAll("p").forEach((p) => {
    const text = (p.textContent || "").trim();
    if (!text && !p.querySelector("img") && !p.querySelector("br")) {
      p.remove();
    }
  });

  // Fix Quill nested list structure safely
  const allLists = tempDiv.querySelectorAll("ul, ol");

  allLists.forEach((list) => {
    const items = Array.from(list.children) as HTMLElement[];
    let currentParentLi: HTMLElement | null = null;
    let nestedList: HTMLElement | null = null;

    items.forEach((item) => {
      if (item.tagName !== "LI") return;

      const li = item as HTMLElement;
      const isIndented =
        li.classList.contains("ql-indent-1") ||
        li.classList.contains("ql-indent-2") ||
        li.classList.contains("ql-indent-3");

      if (!isIndented) {
        currentParentLi = li;
        nestedList = null;
      } else if (currentParentLi) {
        if (!nestedList) {
          nestedList = document.createElement(list.tagName.toLowerCase());
          currentParentLi.appendChild(nestedList);
        }

        li.classList.remove("ql-indent-1", "ql-indent-2", "ql-indent-3");
        nestedList.appendChild(li);
      }
    });
  });

  return tempDiv.innerHTML;
}; 

  return (
    <Modal show={show} onHide={onClose} size="lg" centered>
      <Modal.Header className={styles.modalHeader} closeButton>
        <Modal.Title className={styles.pageTitle}>{item.Title}</Modal.Title>
      </Modal.Header>

      <Modal.Body className={`px-3 py-0 ${styles.modalBody}`}> 
        {/* Carousel */}
        {filteredImages.length > 0 && (
          <Carousel className={styles.modalCarousel}
            interval={3000}
            controls={true}
            indicators={false}
            pause="hover"
            nextIcon={
              <span
                aria-hidden="true"
                className="carousel-control-next-icon"
                style={{ filter: "invert(1)" }}
              />
            }
            prevIcon={
              <span
                aria-hidden="true"
                className="carousel-control-prev-icon"
                style={{ filter: "invert(1)" }}
              />
            }
          >
            {filteredImages.map((img, i) => (
              <Carousel.Item key={i}>
                <img
                  className="d-block w-100"
                  src={img.url}
                  alt="News"
                  style={{
                        height: "auto", 
                    width: "100%",
                    display: "block",
                    margin: "0 auto",
                  }}
                />
                 {img.caption && (
    
      <Carousel.Caption 
      style={{
                  // position: "absolute",
                  // bottom: "0",
                  // left: "0",
                  // width: "100%",
                  background: "rgba(0, 0, 0, 0.55)",
                  padding: "10px 15px",
                  textAlign: "center",
                  color: "white",
                }}>
<p className="mb-0 mt-0">{img.caption}</p>
</Carousel.Caption>  
    )}
              </Carousel.Item>
            ))}
          </Carousel> 
        )}

        {/* Content */}
        <div className="p-1 mt-3">
          <p className="mb-2">
             {formatISTDate(item.EventDate)} 

          </p>

          {/* <h5 className="mb-2">{item.Title}</h5> */}

          {/* MainDescription is HTML (Quill) */}
          {/* <p
            className={`${styles.previewDesc} mt-0 mb-2 `}
            dangerouslySetInnerHTML={{ __html: item.MainDescription || "" }}
          /> */}
  <div
  className={`${styles.previewDesc} mt-0 mb-2`}
  dangerouslySetInnerHTML={{
    __html: normalizeQuillHtml(item.MainDescription || "")
  }}
/>  
        </div>
      </Modal.Body>

      {/* <Modal.Footer>
        <Button variant="outline-primary" onClick={onClose}>
          Close
        </Button>
      </Modal.Footer> */}
    </Modal>
  );
};

export default PreviewNewsModal;
