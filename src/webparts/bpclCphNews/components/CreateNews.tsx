import * as React from "react";
import { useState, useEffect } from "react";
//import "@fontsource/inter"; 
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
import styles from "./MyViewNews.module.scss";
import { Form, Row, Col, Card, Button } from "react-bootstrap";
import * as ReactQuill from "react-quill";
import "bootstrap/dist/css/bootstrap.min.css";
import "react-quill/dist/quill.snow.css";
import PeoplePicker, { Option } from "../services/PeoplePicker";
import PopupModal from "../services/PopupModel";
import { NewsService } from "../services/NewsService";
import ErrorService from "../services/ErrorService";

import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

const QuillEditor = (
  ReactQuill as unknown as {
    default: React.ComponentType<Record<string, unknown>>;
  }
).default;

const modules = {
  toolbar: [
    [{ header: [1, 2, 3, 4, 5, 6, false] }],
    ["bold", "italic", "underline", "strike"],
    [{ color: [] }, { background: [] }],
    [{ script: "sub" }, { script: "super" }],
    [{ list: "ordered" }, { list: "bullet" }],
    [{ indent: "-1" }, { indent: "+1" }],
    [{ align: [] }],
    ["clean"],
  ],
};
 
const formats = [
  "header",
  "bold",
  "italic",
  "underline",
  "strike",
  "color",
  "background",
  "script",
  "list",
  "bullet",
  "indent",
  "align",
];

interface Props {
  context: any;
  form: any;
  setForm: (data: any) => void;
  onBack: () => void;
  onSaveDraft: () => Promise<void>;
  onSubmit: () => Promise<void>;
  onEditAfterSave: (id: number) => void;
  service: NewsService;
  termSetNewsTypesId: string;
  termSetConfidentialityId?: string;
  termSetArchivalPolicyId?: string;

  setThumbnailFile: React.Dispatch<React.SetStateAction<File | null>>;
  setPic1File: React.Dispatch<React.SetStateAction<File | null>>;
  setPic2File: React.Dispatch<React.SetStateAction<File | null>>;
  setPic3File: React.Dispatch<React.SetStateAction<File | null>>;
  attachmentLinks: { FileName: string; ServerRelativeUrl: string }[];
}

const CreateNews: React.FC<Props> = ({
  context,
  service,
  onBack,
  onEditAfterSave,
  termSetNewsTypesId,
  termSetConfidentialityId,
  termSetArchivalPolicyId,
}) => {
  const [title, setTitle] = useState<string>("");
  const [content, setContent] = useState<string>("");
  const [publishIn, setPublishIn] = useState<string>("");
  const [eventDate, setEventDate] = useState<string>("");
  const [approvers, setApprovers] = useState<any[]>([]);
  const [selectedApprover, setSelectedApprover] = useState<string>("");
  const [contentManagers, setContentManagers] = useState<any[]>([]);
  const [selectedContentManager, setSelectedContentManager] =
    useState<string>("");
  const [newsType, setNewsType] = useState<IPickerTerms>([]);
  const [informToBrand, setInformToBrand] = useState<boolean>(false);
  const [addInfo1, setAddInfo1] = useState<string>("");
  const [addInfo2, setAddInfo2] = useState<string>("");
  const [addInfo3, setAddInfo3] = useState<string>("");
  const [comments, setComments] = useState<string>("");
  const [thumbnailFile, setThumbnailFile] = useState<File | null>(null);
  const [thumbnailCaption, setThumbnailCaption] = useState<string>("");
  const [pic1, setPic1] = useState<File | null>(null);
  const [pic2, setPic2] = useState<File | null>(null);
  const [pic3, setPic3] = useState<File | null>(null);

  const [picCap1, setPicCap1] = useState<string>("");
  const [picCap2, setPicCap2] = useState<string>("");
  const [picCap3, setPicCap3] = useState<string>("");

  const [showPopup, setShowPopup] = useState<boolean>(false);
  const [popupMessage, setPopupMessage] = useState<string>("");
  const [popupType, setPopupType] = useState<"success" | "danger" | "confirm">(
    "success",
  );
  const [viewBy, setViewBy] = useState<Option | null>(null);
 
  const [department, setDepartment] = useState<string>("");

  const [savedItemId, setSavedItemId] = useState<number | null>(null);
        const [isSaving, setIsSaving] = useState(false);


  // const [confidentialTerm, setConfidentialTerm] = useState<any>(null);
  // const [archivalTerm, setArchivalTerm] = useState<any>(null); 

  // useEffect(() => {
  //   const loadFixedTerms = async () => {
  //     try {
  //       if (!termSetConfidentialityId || !termSetArchivalPolicyId) return;

  //       const confTerms = await service.getTermSetTerms(
  //         termSetConfidentialityId,
  //       );
  //       const archTerms = await service.getTermSetTerms(
  //         termSetArchivalPolicyId,
  //       );

  //       const conf = confTerms.find(
  //         (t: any) => t.text.toLowerCase() === "confidential",
  //       );

  //       const arch = archTerms.find(
  //         (t: any) => t.text.toLowerCase() === "as per company policy",
  //       );

  //       setConfidentialTerm(conf || null);
  //       setArchivalTerm(arch || null);
  //     } catch (e) {
  //       console.log("Fixed terms load error", e);
  //     }
  //   };

  //   loadFixedTerms();
  // }, [termSetConfidentialityId, termSetArchivalPolicyId]); 
  const confidentialTerm = {
  Label: "Confidential",
  TermGuid: "78e3abf4-3b6a-4fab-ba1c-baca497498cd",
  WssId: -1,
};

const archivalTerm = {
  Label: "As Per Company Policy",
  TermGuid: "93858a26-566e-485e-b8ac-8a480ffcf903",
  WssId: -1,
};

 useEffect(() => {
  const loadData = async () => {
    try {
      //  Get profile (for department display only)
      const profileInfo = await service.getLoggedInUserProfileInfo();
      setDepartment(profileInfo?.department || "");

      //  Validate user (SBU OR Department handled inside)
      const isValid = await service.isUserPartOfSbuTeam();

      if (!isValid) {
        setApprovers([]);
        setContentManagers([]);
        return;
      }

      //  Get approvers (auto handles SBU + Dept + priority)
      let appr = await service.getApproversByTeam();

      //  Add Reporting Manager (if not exists)
      const managerUser = profileInfo?.reportingManagerUser;

      if (managerUser?.Id) {
        const exists = appr.some((u: any) => u.Id === managerUser.Id);
        if (!exists) {
          appr.unshift(managerUser);
        }
      }

      //  Remove current user
      const currentUserId = await service.getCurrentUserId();

      const filteredApprovers = appr.filter(
        (u: any) => u.Id !== currentUserId
      );

      setApprovers(filteredApprovers);

      //  Get Content Managers
      const cms = await service.getContentManagersByTeam();
      setContentManagers(cms);

} catch (error) {

  ErrorService.logError(
    "CreateNews.loadData",
    error
  );

  setPopupType("danger");
  setPopupMessage(
    ErrorService.getUserMessage(error)
  );

  setShowPopup(true);
} 
  }; 

  loadData();
}, []);

  const getExpiryPreview = () => {
    const dt = new Date();
    dt.setDate(dt.getDate() + 90);

    const yyyy = dt.getFullYear();
    const mm = ("0" + (dt.getMonth() + 1)).slice(-2);
    const dd = ("0" + dt.getDate()).slice(-2);

    return `${yyyy}-${mm}-${dd}`;
  };

  

  // validation msg
  const showError = (msg: string) => {
    window.scrollTo({ top: 0, behavior: "smooth" });
    setPopupType("danger");
    setPopupMessage(msg);
    setShowPopup(true);
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

  // saving in the list
  const handleSave = async () => {
          if (isSaving) return;

    if (!title?.trim()) return showError("Title is required");
    
  if (title.trim().length > 255) {
    return showError("Title should not exceed 255 characters.");
  } 
     setIsSaving(true);

    const selectedTerm = newsType?.[0] ?? null;

    const isTransferList =
      selectedTerm?.name?.trim().toLowerCase() === "staff posting";

    const normalizedContent = normalizeQuillHtml(content);


    try {
      //  ONLY ONE CALL
      const result = await service.saveNews(
        title,
        // content,
        normalizedContent,
        publishIn,
        eventDate,
        addInfo1,
        addInfo2,
        addInfo3,
        comments,
        thumbnailFile,
        thumbnailCaption,
        pic1,
        pic2,
        pic3,
        picCap1,
        picCap2,
        picCap3,
        selectedApprover,
        selectedContentManager,
        // newsTypeObj,
        selectedTerm,
        viewBy ? viewBy.value : null,
        informToBrand,
        isTransferList,
        department,
       
    confidentialTerm,
  archivalTerm  

      );

      // Popup with RequestNumber
      setPopupType("success");
      setPopupMessage(
        ` Request Number: ${result.requestNumber} News was saved successfully.`,
      );
      setSavedItemId(result.itemId);

      if (selectedTerm?.name?.trim().toLowerCase() === "staff posting") {
        //await new Promise((res) => setTimeout(res, 1000));
        //await createFolderAndNavigationLinks(result.itemId);
      }

      setShowPopup(true);

      //  Clear Form
      setTitle("");
      setContent("");
      setPublishIn("");
      setEventDate("");
      setAddInfo1("");
      setAddInfo2("");
      setAddInfo3("");
      setComments("");
      setThumbnailFile(null);
      setThumbnailCaption("");
      setPic1(null);
      setPic2(null);
      setPic3(null);
      setPicCap1("");
      setPicCap2("");
      setPicCap3("");
      setSelectedApprover("");
      setSelectedContentManager("");
      //setSelectedNewsTypeKey("");
      setNewsType([]);
      setInformToBrand(false);

      //setNewsType(null);
      setViewBy(null);
  } catch (error) {

  ErrorService.logError(
    "CreateNews.handleSave",
    error
  );

  setPopupType("danger");

  setPopupMessage(
    ErrorService.getUserMessage(error)
  ); 

  setShowPopup(true);
}
finally {
    setIsSaving(false);
  }
   
  };

 
  const validateImageFile = async (
  file: File,
  inputElement: HTMLInputElement,
): Promise<boolean> => {

  // Allowed extensions
  const allowedExtensions = ["bmp", "gif", "png", "jpg", "jpeg"];

  // Allowed MIME types
  const allowedMimeTypes = [
    "image/png",
    "image/jpeg",
    "image/jpg",
    "image/gif",
    "image/bmp",
  ];

  const fileExt = file.name.split(".").pop()?.toLowerCase();

  // Extension validation
  if (!fileExt || !allowedExtensions.includes(fileExt)) {
    setPopupType("danger");
    setPopupMessage(
      "Invalid file extension. Allowed: bmp, gif, png, jpg, jpeg"
    );
    setShowPopup(true);
    inputElement.value = "";
    return false;
  }

  // MIME type validation
  if (!allowedMimeTypes.includes(file.type)) {
    setPopupType("danger");
    setPopupMessage("Invalid file content type.");
    setShowPopup(true);
    inputElement.value = "";
    return false;
  }

  // Size validation (2MB)
  if (file.size > 2 * 1024 * 1024) {
    setPopupType("danger");
    setPopupMessage("Image must be less than 2MB");
    setShowPopup(true);
    inputElement.value = "";
    return false;
  }

  // Magic byte validation
  const buffer = await file.slice(0, 8).arrayBuffer();
  const bytes = new Uint8Array(buffer);

  const isPng =
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4E &&
    bytes[3] === 0x47;

  const isJpg =
    bytes[0] === 0xFF &&
    bytes[1] === 0xD8 &&
    bytes[2] === 0xFF;

  const isGif =
    bytes[0] === 0x47 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46;

  const isBmp =
    bytes[0] === 0x42 &&
    bytes[1] === 0x4D;

  if (!(isPng || isJpg || isGif || isBmp)) {
    setPopupType("danger");
    setPopupMessage("File content does not match image format.");
    setShowPopup(true);
    inputElement.value = "";
    return false;
  }

  return true;
};
 
  // Allow A–Z a–z 0–9 and approved special characters
  //const allowedRegex = /^[A-Za-z0-9()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/; 
const allowedRegex = /^[A-Za-z0-9\u0900-\u097F()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;
  const handleValidatedChange = (
    value: string,
    setter: React.Dispatch<React.SetStateAction<string>>,
  ) => {
    if (allowedRegex.test(value)) {
      setter(value);
    }
  
  }; 

  return (
    <section className={`${styles.createNews} py-2 px-0 px-md-2 py-md-4`}>
      <Card className={`${styles.cardCommon} p-1 p-sm-3 p-md-4`}>
        <Card.Body>
          <h4 className={`${styles.pageTitle}`}>Create News</h4>

          <Form>
            <Row className="g-3">
              <Col xs={12}>
                <Form.Group controlId="title">
  <Form.Label className={styles.floatingLabelStyle}>
    Title <span className={styles.required}>*</span>
  </Form.Label>

  <Form.Control
    className={styles.inputBox}
    type="text"
    placeholder="Enter title"
    value={title}
    maxLength={255}
    onChange={(e) =>
      handleValidatedChange(e.target.value, setTitle)
    }
    required
  /> 

  <small style={{ color: title.length > 250 ? "red" : "#666" }}>
    {title.length}/255 characters
  </small>
</Form.Group>  
              </Col>

              <Col xs={12}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Description <span className={styles.required}>*</span>
                  </Form.Label>
                  <div className={styles.quillWrapper}>
                    <QuillEditor
                      theme="snow"
                      value={content}
                      onChange={setContent}
                      modules={modules}
                      formats={formats}
                      placeholder="Enter news content"
                    />
                  </div>
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="category" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Publish In <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    required
                    value={publishIn}
                    onChange={(e) => setPublishIn(e.target.value)}
                  >
                    <option value="">Select</option>
                    <option value="corporate">Corporate</option>
                    <option value="sbu">My Team</option>
                  </Form.Select>
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    News Type <span className={styles.required}>*</span>
                  </Form.Label>

                  <TaxonomyPicker
                    context={context as any}
                    allowMultipleSelections={false}
                    label=""
                    termsetNameOrID="7b5de19a-32f9-4e96-bcc4-ede23914510b"
                    anchorId="5cf3133c-23d5-4581-9d77-5ed7e212a7a5"
                    panelTitle="Select News Type"
                    onChange={(terms) => setNewsType(terms || [])}
                    isTermSetSelectable={false}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={12} md={4}>
                <Form.Group controlId="confidentialityIndex" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Confidentiality Index{" "}
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={"Confidential"}
                    readOnly
                  />
                </Form.Group>
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="archivalPolicy" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Archival Policy
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={"As Per Company Policy"}
                    readOnly
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="eventDate" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Event Date <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="date"
                    placeholder="Select event date"
                    required
                    value={eventDate}
                    onChange={(e) => setEventDate(e.target.value)}
                  />
                </Form.Group>
              </Col>
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="expiryDate" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Expiry Date
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="date"
                    value={getExpiryPreview()}
                    readOnly
                  />
                </Form.Group>
              </Col>
            </Row>

            <Row className=" mt-1 g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="approver" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Approver <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    required
                    value={selectedApprover}
                    onChange={(e) => setSelectedApprover(e.target.value)}
                  >
                    <option value="">Select</option>
                    {approvers.map((u) => (
                      <option key={u.Id} value={u.Id}>
                        {u.Title}
                      </option>
                    ))}
                  </Form.Select>
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="contentManager" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Content Manager <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    required
                    value={selectedContentManager}
                    onChange={(e) => setSelectedContentManager(e.target.value)}
                  >
                    <option value="">Select</option>
                    {contentManagers.map((u) => (
                      <option key={u.Id} value={u.Id}>
                        {u.Title}
                      </option>
                    ))}
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* View By */}
              <Col xs={12} sm={6} md={4}>
                <Form.Label className={styles.floatingLabelStyle}>
                  View By
                </Form.Label>
                <PeoplePicker
                  context={context}
                  isClearable={true}
                  selectedValue={viewBy}
                  onChange={(val) => setViewBy(val)}
                />
              </Col>
            </Row>

            <Row className="mt-2">
              <Col>
                <Form.Check
                  inline
                  type="checkbox"
                  id="informToBrand"
                  label="Inform To Brand (would you like brand team to take your publications)"
                  checked={informToBrand}
                  onChange={(e) => setInformToBrand(e.target.checked)}
                />
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              <Col xs={12} md={6}>
                <Form.Group controlId="thumbnail">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Thumbnail <span className={styles.required}>*</span>
                  </Form.Label>
                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    required
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                   onChange={async (
  e: React.ChangeEvent<HTMLInputElement>,
) => {

  const input = e.target;

  const file = input.files?.[0];

  if (!file) return;

  const isValid = await validateImageFile(
    file,
    input
  );

  if (!isValid) return;

  setThumbnailFile(file);
}}
                  />
                  <small>
                    (Image should be less than 2MB and allowed bmp, gif, png,
                    jpg, jpeg)
                  </small>
                </Form.Group>
              </Col>

              <Col xs={12} md={6}>
                <Form.Group controlId="thumbnailCaption">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Thumbnail Caption
                  </Form.Label>
                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Thumbnail caption"
                    value={thumbnailCaption}
                    onChange={(e) =>
                      handleValidatedChange(e.target.value, setThumbnailCaption)
                    }
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Pictures */}
            <Row className="mt-1 g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture1">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 1
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    onChange={async (e) => {

  const input = e.target as HTMLInputElement;

  const file = input.files?.[0];

  if (!file) return;

  const isValid = await validateImageFile(
    file,
    input
  );

  if (!isValid) return;

  setPic1(file);
}}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture2">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 2
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    onChange={async (e) => {

  const input = e.target as HTMLInputElement;

  const file = input.files?.[0];

  if (!file) return;

  const isValid = await validateImageFile(
    file,
    input
  );

  if (!isValid) return;

  setPic2(file);
}}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture3">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 3
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    onChange={async (e) => {

  const input = e.target as HTMLInputElement;

  const file = input.files?.[0];

  if (!file) return;

  const isValid = await validateImageFile(
    file,
    input
  );

  if (!isValid) return;

  setPic3(file);
}}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Picture Captions */}
            <Row className="mt-1  g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pictureCaption1" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 1 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Picture 1 Caption"
                    value={picCap1}
                    onChange={(e) =>
                      handleValidatedChange(e.target.value, setPicCap1)
                    }
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pictureCaption2" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 2 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Picture 2 Caption"
                    value={picCap2}
                    onChange={(e) =>
                      handleValidatedChange(e.target.value, setPicCap2)
                    }
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pictureCaption3" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 3 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Picture 3 Caption"
                    value={picCap3}
                    onChange={(e) =>
                      handleValidatedChange(e.target.value, setPicCap3)
                    }
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Additional Comments */}
            <Row className="mt-1 g-3">
              {["1", "2", "3"].map((num, index) => (
                <Col key={num} xs={12} sm={6} md={4}>
                  <Form.Group controlId={`additionalInfo${num}`} className="">
                    <Form.Label className={styles.floatingLabelStyle}>
                      {`Additional Info ${num}`}
                    </Form.Label>

                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      placeholder={`Additional Info ${num}`}
                      value={
                        index === 0
                          ? addInfo1
                          : index === 1
                            ? addInfo2
                            : addInfo3
                      }
                      onChange={(e) => {
                        const value = e.target.value;

                        if (!allowedRegex.test(value)) return; // silently block

                        if (index === 0) setAddInfo1(value);
                        if (index === 1) setAddInfo2(value);
                        if (index === 2) setAddInfo3(value);
                      }}
                    />
                  </Form.Group>
                </Col>
              ))}
            </Row>

            <Row className="mt-1 g-3">
              <Col>
                <Form.Group controlId="comment" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Comment
                  </Form.Label>
                  <Form.Control
                    as="textarea"
                    placeholder="Comment"
                    value={comments}
                    style={{
                      minHeight: "50px",
                      resize: "none",
                      overflow: "hidden",
                    }}
                    onChange={(e) => {
                      e.target.style.height = "auto";
                      e.target.style.height = `${Math.max(e.target.scrollHeight, 50)}px`;
                      const value = e.target.value;
                      if (allowedRegex.test(value)) {
                        setComments(value);
                      }
                    }}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Buttons */}
            <Row className="mt-4">
              <div className="d-flex align-items-center w-100">
                <div className="ms-auto d-flex gap-2">
                  <Button
                    className={`${styles.primaryOutlineBtn} btn`}
                    onClick={onBack}
                  >
                    Cancel
                  </Button>

                  <Button
                    className={`${styles.primaryBtn} btn`}
                    onClick={handleSave}
                   disabled={isSaving}
>
  {isSaving ? "Saving..." : "Save"}
</Button>
                </div>
              </div>
            </Row>
          </Form>
        </Card.Body>
      </Card>
      <PopupModal
        show={showPopup}
        type={popupType}
        message={popupMessage}
        onClose={() => {
          setShowPopup(false);

          //  After popup OK -> navigate to edit
          if (popupType === "success" && savedItemId) {
            onEditAfterSave(savedItemId);
          }
        }}
      />
    </section>
  );
};

export default CreateNews;
