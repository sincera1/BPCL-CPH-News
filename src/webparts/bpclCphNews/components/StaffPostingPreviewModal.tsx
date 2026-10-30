import * as React from "react";
import { useEffect, useState } from "react";
import { Modal, Form, Carousel } from "react-bootstrap";
import { NewsService } from "../services/NewsService";
import styles from "./MyViewNews.module.scss";
import ErrorService from "../services/ErrorService";
 
 

interface Props {
  show: boolean;
  onClose: () => void;
  corpItemId: number;
  service: NewsService;
}

const StaffPostingPreviewModal: React.FC<Props> = ({
  show,
  onClose,
  corpItemId,
  service,
}) => {
  const [sbuOptions, setSbuOptions] = useState<any[]>([]);
  const [selectedSbu, setSelectedSbu] = useState<string>("");
  const [selectedStartIndex, setSelectedStartIndex] = useState<string>("");
  const [images, setImages] = useState<any[]>([]);
  const [filteredImages, setFilteredImages] = useState<any[]>([]);

  useEffect(() => {
    if (!show) return;

    const load = async () => {
      const sbuItems = await service.getTransferSbuIndexesByCorpId(corpItemId);
      const list = sbuItems || [];

      setSbuOptions(list);

      if (list.length === 0) {
        setSelectedSbu("");
        setSelectedStartIndex("");
        return;
      }

      //  Priority: Office Note first
      const officeNoteRow = list.find(
        (x: any) => (x.SBU?.TeamName || "").toLowerCase() === "office note",
      );

      const defaultRow = officeNoteRow || list[0];

      setSelectedSbu(defaultRow.SBU?.TeamName || "");
      setSelectedStartIndex(defaultRow.StartIndex?.toString() || "");
    };

    load();
  }, [show, corpItemId, service]);

  //  load images from folder
  useEffect(() => {
    if (!show) return;

    let interval: any;

    const loadImages = async () => {
      try {
        const files = await service.getTransferImagesByCorpId(corpItemId);
        setImages(files || []);
          } catch (error) {
    ErrorService.logError(
      "StaffPostingPreviewModal.loadImages",
      error
    ); 
        setImages([]);
      }
    }; 

    loadImages(); //  first time load

    // auto refresh every 3 seconds
    interval = setInterval(loadImages, 3000);

    return () => clearInterval(interval);
  }, [show, corpItemId]);

useEffect(() => {
  if (!selectedStartIndex || images.length === 0) {
    setFilteredImages([]);
    return;
  } 

  const start = parseFloat(selectedStartIndex);

  const getNum = (name: string) => {
    const match = name?.toLowerCase().match(/^(\d+(\.\d+)?)/);
    return match ? parseFloat(match[1]) : 999999;
  };

  // Step 1: sort images
  const sorted = [...images].sort(
    (a, b) => getNum(a.Name) - getNum(b.Name)
  );

  // Step 2: find start position
  const startIndex = sorted.findIndex(
    (img) => getNum(img.Name) >= start
  );

  if (startIndex === -1) {
    setFilteredImages(sorted);
    return;
  }

  // Step 3: rotate images
  const rotated = [
    ...sorted.slice(startIndex),
    ...sorted.slice(0, startIndex),
  ];

  setFilteredImages(rotated);

}, [selectedStartIndex, images]);  

    
  return (
    <Modal show={show} onHide={onClose} size="lg" centered>
      <Modal.Header className={styles.modalHeader} closeButton>
        <Modal.Title>Staff Posting Preview</Modal.Title>
      </Modal.Header>

      <Modal.Body className={`px-3 py-0 ${styles.modalBody}`}> 
        <div className="mb-3">
          <Form.Label>Select SBU</Form.Label>

          <Form.Select
            value={selectedSbu}
            onChange={(e) => {
              const sbu = e.target.value;
              setSelectedSbu(sbu);

              const row = sbuOptions.find((x: any) => x.SBU?.TeamName === sbu);

              setSelectedStartIndex(row?.StartIndex?.toString() || "");
            }}
          >
            {sbuOptions.map((x: any) => (
              <option key={x.Id} value={x.SBU?.TeamName || ""}>
                {x.SBU?.TeamName}
              </option>
            ))}
          </Form.Select>
        </div>
 
        {/*  images carousel */}
        <div>
          {filteredImages.length === 0 ? (
            <div className="text-muted text-center p-3">
              No images found {selectedStartIndex || ""} 
            </div>
          ) : (
            <Carousel
              interval={3000}
              controls={true}
              indicators={true}
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
              {filteredImages.map((img: any, i: number) => {
                const imgUrl =
                  img.ServerRelativeUrl ||
                  img.ServerRelativePath?.DecodedUrl ||
                  "";
 
                return (
                  <Carousel.Item key={imgUrl || i}>
                    <img
                      className="d-block w-100"
                      src={`${imgUrl}?t=${Date.now()}`}
                      alt={`Staff Posting ${i + 1}`}
                      style={{
                        height: "auto",
                        width: "100%",
                        objectFit: "scale-down",
                        display: "block",
                        margin: "0 auto",
                        borderRadius: "8px",
                      }}
                    />
                  </Carousel.Item>
                );
              })}
            </Carousel>
          )}
        </div>
      </Modal.Body>
    </Modal> 
  );
};

export default StaffPostingPreviewModal;
