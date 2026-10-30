import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { Form, Row, Col, Card, Button } from "react-bootstrap";
//import "@fontsource/inter"; // default (400)
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";

import { SPHttpClient } from "@microsoft/sp-http";

import { TaxonomyPicker } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

import styles from "./MyViewNews.module.scss";
import { NewsService } from "../services/NewsService";
import PopupModal from "../services/PopupModel";

import ErrorService from "../services/ErrorService";

import PeoplePickerCustom, { Option } from "../services/PeoplePicker";

import * as ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";

const QuillEditor = (ReactQuill as any).default;

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

  editId: number | null;

  onBack: () => void;

  refreshList?: () => Promise<void>;

  onSaveDraft: () => Promise<void>;
  onSubmit: () => Promise<void>;
  onPreviewPopup: () => void;

  service: NewsService;

  termSetNewsTypesId: string;
  termSetConfidentialityId?: string;
  termSetArchivalPolicyId?: string;

  setThumbnailFile: React.Dispatch<React.SetStateAction<File | null>>;
  setPic1File: React.Dispatch<React.SetStateAction<File | null>>;
  setPic2File: React.Dispatch<React.SetStateAction<File | null>>;
  setPic3File: React.Dispatch<React.SetStateAction<File | null>>;

  thumbnailFile: File | null;
  pic1File: File | null;
  pic2File: File | null;
  pic3File: File | null;

  attachmentLinks: { FileName: string; ServerRelativeUrl: string }[];
}

const NewsEditForm: React.FC<Props> = ({
  context,
  form,
  setForm,
  editId,
  onBack,

  refreshList,

  onSaveDraft,
  onSubmit,
  termSetNewsTypesId,
  termSetConfidentialityId,
  termSetArchivalPolicyId,
  attachmentLinks,
  service,
  onPreviewPopup,
  setThumbnailFile,
  setPic1File,
  setPic2File,
  setPic3File,

  thumbnailFile,
  pic1File,
  pic2File,
  pic3File,
}) => {
  // dropdown data like CreateNews
  const [approvers, setApprovers] = useState<any[]>([]);

  const [contentManagers, setContentManagers] = useState<any[]>([]);
  const [selectedContentManager, setSelectedContentManager] =
    useState<string>("");

  // ViewBy custom picker
  const [viewBy, setViewBy] = useState<Option | null>(null);

  const [newsType, setNewsType] = useState<IPickerTerms>([]);

  const [isNewsCreator, setIsNewsCreator] = useState(false);
  const [isApprover, setIsApprover] = useState(false);
  const [isPrimaryCM, setIsPrimaryCM] = useState(false);
  const [currentUserId, setCurrentUserId] = useState<number>(0);

  const [isStaffPosting, setIsStaffPosting] = useState<boolean>(false);
  //const [transferListUrl, setTransferListUrl] = useState<string>("");
  const [imagesFolderUrl, setImagesFolderUrl] = useState<string>("");

  const [confidentialTerm, setConfidentialTerm] = useState<any>(null);
  const [archivalTerm, setArchivalTerm] = useState<any>(null);

  const commentRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    const adjustHeight = (textarea: HTMLTextAreaElement | null) => {
      if (textarea) {
        textarea.style.height = "auto";
        textarea.style.height = textarea.scrollHeight + "px";
      }
    };

    adjustHeight(commentRef.current);
  }, []);

  useEffect(() => {
    const loadFixedTerms = async () => {
      if (!termSetConfidentialityId || !termSetArchivalPolicyId) return;

      const confTerms = await service.getTermSetTerms(termSetConfidentialityId);
      const archTerms = await service.getTermSetTerms(termSetArchivalPolicyId);

      const conf = confTerms.find(
        (t: any) => t.text.toLowerCase() === "confidential",
      );

      const arch = archTerms.find(
        (t: any) => t.text.toLowerCase() === "as per company policy",
      );

      setConfidentialTerm(conf || null);
      setArchivalTerm(arch || null);
    };

    loadFixedTerms();
  }, [termSetConfidentialityId, termSetArchivalPolicyId]);

  useEffect(() => {
    if (!confidentialTerm || !archivalTerm) return;

    setForm((prev: any) => ({
      ...prev,
      ConfidentialityIndex: {
        Label: confidentialTerm.text,
        TermGuid: confidentialTerm.key,
      },
      ArchivalPolicy: {
        Label: archivalTerm.text,
        TermGuid: archivalTerm.key,
      },
    }));
  }, [confidentialTerm, archivalTerm]);

  const [showPopup, setShowPopup] = useState(false);
  const [popupType, setPopupType] = useState<"success" | "danger" | "confirm">(
    "success",
  );
  const [popupMessage, setPopupMessage] = useState("");
  const [confirmAction, setConfirmAction] = useState<
    (() => Promise<void>) | null
  >(null);

  // term name dislpay
  const getTermNameByGuid = async (guid: string): Promise<string> => {
    try {
      const items = await service.sp.web.lists
        .getByTitle("TaxonomyHiddenList")
        .items.filter(`IdForTerm eq '${guid}'`)
        .select("Term")
        .top(1)(); 

      return items?.[0]?.Term || "";
    } catch (e) {
      return "";
    }
  };

  useEffect(() => {
    const loadNewsTypeTerm = async () => {
      if (form?.NewsTypes?.TermGuid && termSetNewsTypesId) {
        const termName = await getTermNameByGuid(form.NewsTypes.TermGuid);

        if (!termName) return;

        //  Set picker selected value
        setNewsType([
          {
            key: form.NewsTypes.TermGuid,
            name: termName, // REAL TERM LABEL
            path: `${termSetNewsTypesId}|${termName}`,
            termSet: termSetNewsTypesId,
          },
        ]);

        //  Also update form so save uses correct label
        setForm((prev: any) => ({
          ...prev,
          NewsTypes: {
            Label: termName,
            TermGuid: form.NewsTypes.TermGuid,
          },
        }));
      } else {
        setNewsType([]);
      }
    };

    loadNewsTypeTerm();
  }, [form?.NewsTypes?.TermGuid, termSetNewsTypesId]);

  // links checking nagivation urls list and library.
  useEffect(() => {
    const loadTransferUrls = async () => {
      if (!editId) return;

      //  Get selected taxonomy label from form
      const selectedLabel = form?.NewsTypes?.Label?.trim().toLowerCase() || "";

      if (selectedLabel === "staff posting") {
        setIsStaffPosting(true);

        const webUrl = context.pageContext.web.absoluteUrl;

        // ===== LIBRARY FOLDER URL =====
        const libFolderPath = `${context.pageContext.web.serverRelativeUrl}/TransferListImages/${editId}`;

        setImagesFolderUrl(
          `${webUrl}/TransferListImages/Forms/AllItems.aspx?RootFolder=${encodeURIComponent(
            libFolderPath,
          )}`,
        );
      } else {
        setIsStaffPosting(false);
        setImagesFolderUrl("");
      }
    };

    loadTransferUrls();
  }, [editId, form?.NewsTypes]);

  // creating folder transfersbuindexlist
  const openTransferFolder = async (itemId: number) => {
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const listTitle = "TransferSbuIndexList";
      const folderTitle = itemId.toString();

      // 1 Get folder by Title (visible name) to check exist
      const response = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$filter=Title eq '${folderTitle}' and FSObjType eq 1&$select=FileRef`,
        SPHttpClient.configurations.v1,
      );

      const data = await response.json();

      if (data.value && data.value.length > 0) {
        //  Folder found
        const folderPath = data.value[0].FileRef;

        const folderUrl = `${webUrl}/Lists/${listTitle}/AllItems.aspx?RootFolder=${encodeURIComponent(folderPath)}`;

        window.open(folderUrl, "_blank");
        return;
      }

      //  Folder not found → create it
      // 1️ Get entity type
      const typeResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getByTitle('${listTitle}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
      );

      if (!typeResponse.ok) {
ErrorService.logError(
  "NewsEditForm.openTransferFolder",
  {
    message: "Failed to get entity type"
  }
);        return;
      }

      const typeData = await typeResponse.json();
      const entityType = typeData.ListItemEntityTypeFullName;

      // 3️ Create folder
      const createResponse = await context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items`,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({
            "@odata.type": entityType,
            Title: folderTitle,
            FileLeafRef: folderTitle,
            ContentTypeId: "0x0120",
          }),
        },
      );

      if (!createResponse.ok) {
        const error = await createResponse.text();
ErrorService.logError(
  "NewsEditForm.openTransferFolder",
  error
);        return;
      }

      // 3️ After create, fetch FileRef

      const newFolderResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$filter=Title eq '${folderTitle}' and FSObjType eq 1&$select=FileRef`,
        SPHttpClient.configurations.v1,
      );

      const newData = await newFolderResponse.json();

      if (newData.value && newData.value.length > 0) {
        const folderPath = newData.value[0].FileRef;

        const folderUrl = `${webUrl}/Lists/${listTitle}/AllItems.aspx?RootFolder=${encodeURIComponent(folderPath)}`;

        window.open(folderUrl, "_blank");
      }
    } catch (err) {
ErrorService.logError(
  "NewsEditForm.openTransferFolder",
  err
);    }
  };

  // creating folder transferlist images library

  const openImagesLibraryFolder = async (itemId: number) => {
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const serverUrl = context.pageContext.web.serverRelativeUrl;

      const libraryName = "TransferListImages";
      const folderName = itemId.toString();

      const folderServerPath = `${serverUrl}/${libraryName}/${folderName}`;

      /* ---------- CHECK FOLDER EXISTS ---------- */

      const checkResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${folderServerPath}')`,
        SPHttpClient.configurations.v1,
      );

      if (!checkResponse.ok) {
        /* ---------- CREATE FOLDER ---------- */
        await service.createFolderIfNotExists(libraryName, folderName);
      }

      /* ---------- OPEN FOLDER ---------- */

      const folderUrl = `${webUrl}/${libraryName}/Forms/AllItems.aspx?RootFolder=${encodeURIComponent(folderServerPath)}`;

      window.open(folderUrl, "_blank");
    } catch (err) {
ErrorService.logError(
  "NewsEditForm.openImagesLibraryFolder",
  err
);    }
  };

 
 useEffect(() => {
  const loadPeopleData = async () => {
    try {
      //  Get profile (only for UI use)
      const profileInfo = await service.getLoggedInUserProfileInfo();
      const managerUser = profileInfo?.reportingManagerUser;

      //  Validate mapping (SBU + Department handled inside)
//       const isValid = await service.isUserPartOfSbuTeam();

//       if (!isValid) {
// ErrorService.logError(
//   "NewsEditForm.loadPeopleData",
//   {
//     message: "Something went wrong. Please contact administrator."
//   }
// );        
//         return;
//       }

      //  Current user
      const currentUserId = await service.getCurrentUserId();
      setCurrentUserId(currentUserId);

      //  Roles
      const creator = await service.isCurrentUserNewsCreator();
      setIsNewsCreator(creator);

      const cm = await service.isCurrentUserPrimaryContentManager();
      setIsPrimaryCM(cm);

      const approverFromTeam = await service.isCurrentUserApprover();

      const isAssignedApprover = form.AssignTo?.[0] === currentUserId;
      setIsApprover(approverFromTeam || isAssignedApprover);

      //  Approvers dropdown
      let appr = await service.getApproversByTeam();

      // Add manager if not exists
      if (managerUser?.Id) {
        const exists = appr.some((u: any) => u.Id === managerUser.Id);
        if (!exists) {
          appr.unshift(managerUser);
        }
      }

      // Remove creator
      const itemAuthorId = form?.AuthorId || null;
      let finalApprovers = appr.filter((u: any) => u.Id !== itemAuthorId);

      // Ensure assigned approver exists
      const pendingId = form?.AssignTo?.[0];

      if (pendingId) {
        const exists = finalApprovers.some((u: any) => u.Id === pendingId);

        if (!exists) {
          try {
            const pendingUser =
              await service.sp.web.siteUsers.getById(pendingId)();
            finalApprovers.unshift(pendingUser);
          } catch (error) {
ErrorService.logError(
  "NewsEditForm.loadPeopleData.pendingUser",
  error
);          }
        }
      }

      setApprovers(finalApprovers);

      //  Content Managers dropdown
      const cms = await service.getContentManagersByTeam();
      setContentManagers(cms);

      if (form.FollowOnActionBy?.length > 0) {
        setSelectedContentManager(String(form.FollowOnActionBy[0]));
      }

      //  ViewBy
      if (form.ViewBy?.length > 0) {
        const id = Number(form.ViewBy[0]);

        try {
          const user = await service.sp.web.siteUsers.getById(id)();

          setViewBy({
            label: `${user.Title}${user.Email ? ` (${user.Email})` : ""}`,
            value: user.Id,
            type: "User",
            title: user.Title,
            email: user.Email,
            loginName: user.LoginName,
          });
        } catch {
          try {
            const group = await service.sp.web.siteGroups.getById(id)();

            setViewBy({
              label: `${group.Title} (Group)`,
              value: group.Id,
              type: "Group",
              title: group.Title,
            });
          } catch (error) {
ErrorService.logError(
  "NewsEditForm.loadPeopleData.viewBy",
  error
);            setViewBy(null);
          }
        }
      }
    } catch (error) {
ErrorService.logError(
  "NewsEditForm.loadPeopleData",
  error
);

showActionPopup(
  ErrorService.getUserMessage(error),
  "danger"
); 
    }
  };
 
  loadPeopleData();
}, [form?.AuthorId, form?.AssignTo, form?.FollowOnActionBy, form?.Status]);

  // const pendingWithId = form.PendingWith?.[0] || null;
  // const isPendingWithMe = pendingWithId === currentUserId;

  const pendingWithId = Number(form.PendingWith?.[0] || 0);
const isPendingWithMe =
  pendingWithId > 0 &&
  pendingWithId === Number(currentUserId);

  const assignedtoId = form.AssignTo?.[0] || null;

  const FollowOnActionById = form.FollowOnActionBy?.[0] || null;

  const currentStatus = (form.Status || "").toString().trim().toLowerCase();

  

  // ---------------- STATUS ----------------

  const isDraft = currentStatus === "draft";
  const isPendingCreator = currentStatus === "pendingwithcreator";
  const isPendingLM = currentStatus === "pendingwithlm";
  const isPendingCM = currentStatus === "pendingwithcm";

  const isFinalStatus =
    currentStatus === "rejected" || currentStatus === "published";

  // ---------------- CREATOR ----------------
  const canCreatorEdit =
    isNewsCreator &&
    form.AuthorId === currentUserId &&
    (isDraft || isPendingCreator);

  // ---------------- LM ----------------

  // Approver: TeamCMList OR direct assignment
const canActAsApprover =
  isApprover || isPendingWithMe;

// Line Manager permissions
const canShowLMButtons =
  isPendingLM &&
  isPendingWithMe &&
  canActAsApprover;
  
 // const canShowLMButtons = isPendingLM && isApprover && isPendingWithMe;

  // ---------------- CM ----------------
  const canShowCMButtons = isPendingCM && isPrimaryCM && isPendingWithMe;

  // ---------------- SUBMIT ----------------
  // CM-only who created this item
  const isCMOwnDraft =
    isPrimaryCM && form.AuthorId === currentUserId && isDraft;

  // Final Submit rule
  const canShowSubmitButton =
    !isFinalStatus && (canCreatorEdit || isCMOwnDraft);

  // ---------------- SAVE ----------------

  const canShowSaveButton =
    !isFinalStatus && (canCreatorEdit || canShowLMButtons || isCMOwnDraft);

  // ---------------- DELETE ----------------
  // Creator OR LM
  const canShowDeleteButton =
    !isFinalStatus && (canCreatorEdit || canShowLMButtons || isCMOwnDraft);

  // ---------------- APPROVE / REJECT ----------------
  // LM + CM
  const canShowApproveRejectButtons =
    !isFinalStatus && (canShowLMButtons || canShowCMButtons);

  // ---------------- REASSIGN ----------------
  // LM + CM
  const canShowReassignButton =
    !isFinalStatus && (canShowLMButtons || canShowCMButtons);

  // Edit the fields
  const canEditFields = canCreatorEdit || canShowLMButtons;

  const disableWorkflowDropdowns = !canEditFields || canShowLMButtons;

  // ---------------- COMMENT EDIT ----------------

  // CM can edit comment only when PendingWithCM and assigned to them
  const canCMEditComment =
    isPrimaryCM && isPendingCM && isPendingWithMe && !isFinalStatus;

  // Final comment permission
  const canEditComment = canEditFields || canCMEditComment;


  const formatRequestNumber = (id: number | null) => {
    if (!id) return "-";
    const padded = ("00000000" + id.toString()).slice(-8);
    return `RQ${padded}`;
  };

  const showActionPopup = (
    msg: string,
    type: "success" | "danger" = "success",
  ) => {
    setPopupType(type);
    setPopupMessage(`Request Number: ${formatRequestNumber(editId)} - ${msg}`);
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
  // Approve
  const onApprove = async () => {
    if (!editId) return;

    try {
     
      // LM / Approver
// if (isApprover && isPendingLM && isPendingWithMe) {
if (canShowLMButtons) {
  const approverId = form.AssignTo?.[0];
  const cmId = form.FollowOnActionBy?.[0];

  const isSamePerson = approverId && cmId && approverId === cmId;

  //  CASE 1: Approver and CM are SAME → Direct Publish
  if (isSamePerson) {

    await service.updateNews(editId, {
      Title: form.Title,
MainDescription: normalizeQuillHtml(form.MainDescription),
      PublishIn: form.PublishIn,

      NewsTypes: form.NewsTypes
        ? service.getTaxonomyValue(form.NewsTypes)
        : null,

      // ConfidentialityIndex: form.ConfidentialityIndex
      //   ? service.getTaxonomyValue(form.ConfidentialityIndex)
      //   : null,

      // ArchivalPolicy: form.ArchivalPolicy
      //   ? service.getTaxonomyValue(form.ArchivalPolicy)
      //   : null,

      EventDate: form.EventDate
        ? new Date(form.EventDate).toISOString()
        : null,

      AssignToId: form.AssignTo?.[0] || null,
      FollowOnActionById: form.FollowOnActionBy?.[0] || null,
      ViewById: form.ViewBy?.[0] || null,

      InformToBrand: form.InformToBrand,

      Thumbnail: form.Thumbnail,
      ThumbnailCaption: form.ThumbnailCaption,

      Picture1: form.Picture1,
      Pic1Caption: form.Pic1Caption,

      Picture2: form.Picture2,
      Pic2Caption: form.Pic2Caption,

      Picture3: form.Picture3,
      Pic3Caption: form.Pic3Caption,

      AdditionalInformation1: form.AdditionalInformation1,
      AdditionalInformation2: form.AdditionalInformation2,
      AdditionalInformation3: form.AdditionalInformation3,

      Comments: form.Comments,

      // DIRECT PUBLISH
      Status: "Published",
      PublishedDate: new Date(),
      PendingWithId: null 
    });

       /* ================= WAIT FOR SP ================= */
        await new Promise((r) => setTimeout(r, 600));

        /* ================= WAIT FOR SP ================= */

        const files: File[] = [];

        if (thumbnailFile) files.push(thumbnailFile);
        if (pic1File) files.push(pic1File);
        if (pic2File) files.push(pic2File);
        if (pic3File) files.push(pic3File);

        if (files.length > 0) {
          await service.uploadAttachments(editId, files);
        }

        await new Promise((r) => setTimeout(r, 800));

        const expiry = await service.getExpiryDateById(editId);

        if (expiry) {
          setForm((prev: any) => ({
            ...prev,
            ExpiryDate: expiry.split("T")[0],
          }));
        }

        /* clear states */
        setThumbnailFile(null);
        setPic1File(null);
        setPic2File(null);
        setPic3File(null);
  

    showActionPopup("Published Successfully");
    return;
  }

  //  CASE 2: Normal Flow (LM → CM)
  await service.updateNews(editId, {
    Title: form.Title,
MainDescription: normalizeQuillHtml(form.MainDescription),
    PublishIn: form.PublishIn,

    NewsTypes: form.NewsTypes
      ? service.getTaxonomyValue(form.NewsTypes)
      : null,

    // ConfidentialityIndex: form.ConfidentialityIndex
    //   ? service.getTaxonomyValue(form.ConfidentialityIndex)
    //   : null,

    // ArchivalPolicy: form.ArchivalPolicy
    //   ? service.getTaxonomyValue(form.ArchivalPolicy)
    //   : null,

    EventDate: form.EventDate
      ? new Date(form.EventDate).toISOString()
      : null,

    AssignToId: form.AssignTo?.[0] || null,
    FollowOnActionById: form.FollowOnActionBy?.[0] || null,
    ViewById: form.ViewBy?.[0] || null,

    InformToBrand: form.InformToBrand,

    Thumbnail: form.Thumbnail,
    ThumbnailCaption: form.ThumbnailCaption,

    Picture1: form.Picture1,
    Pic1Caption: form.Pic1Caption,

    Picture2: form.Picture2,
    Pic2Caption: form.Pic2Caption,

    Picture3: form.Picture3,
    Pic3Caption: form.Pic3Caption,

    AdditionalInformation1: form.AdditionalInformation1,
    AdditionalInformation2: form.AdditionalInformation2,
    AdditionalInformation3: form.AdditionalInformation3,

    Comments: form.Comments,

    Status: "PendingWithCM",
    PendingWithId: FollowOnActionById ? Number(FollowOnActionById) : null 
  });

     /* ================= WAIT FOR SP ================= */
        await new Promise((r) => setTimeout(r, 600));

        /* ================= WAIT FOR SP ================= */

        const files: File[] = [];

        if (thumbnailFile) files.push(thumbnailFile);
        if (pic1File) files.push(pic1File);
        if (pic2File) files.push(pic2File);
        if (pic3File) files.push(pic3File);

        if (files.length > 0) {
          await service.uploadAttachments(editId, files);
        }

        await new Promise((r) => setTimeout(r, 800));

        const expiry = await service.getExpiryDateById(editId);

        if (expiry) {
          setForm((prev: any) => ({
            ...prev,
            ExpiryDate: expiry.split("T")[0],
          }));
        }

        /* clear states */
        setThumbnailFile(null);
        setPic1File(null);
        setPic2File(null);
        setPic3File(null); 


  showActionPopup("Approved Successfully");
  return;
} 

      // CM → Publish
      if (isPrimaryCM && isPendingCM && isPendingWithMe) {
        await service.updateNews(editId, {
          Status: "Published",
          PublishedDate: new Date(),

          PendingWithId: null,
        });

        showActionPopup("Published Successfully");

        return;
      }

      
    } catch (error) {

  ErrorService.logError(
    "NewsEditForm.onApprove",
    error
  );

 showActionPopup(
  ErrorService.getUserMessage(error),
  "danger"
);


}
  };

  // Reject
  const onReject = async () => {
    if (!form.Comments || !form.Comments.trim()) {
      setPopupType("danger");
      setPopupMessage("Comment is required before this action.");
      setShowPopup(true);
      return;
    }
    if (!editId) return;
    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} - Are you sure you want to Reject this News?`,
    );

    setConfirmAction(() => async () => {
      try {
        await service.updateNews(editId, {
          Status: "Rejected",
          PendingWithId: null,
          Comments: form.Comments,
        });
        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(editId)} - Rejected Successfully.`,
        );
        setShowPopup(true); // Popup will handle refresh + back
      } catch (error) {
ErrorService.logError(
  "NewsEditForm.onReject",
  error
);        setPopupType("danger");
       setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
);
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  };

  // Delete
  const onDelete = async () => {
    if (!editId) return;

    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(editId)} - Are you sure you want to delete this News permanently?`,
    );

    setConfirmAction(() => async () => {
      try {
        await service.deleteNewsPermanently(editId);

        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(editId)} - Deleted permanently.`,
        );
        setShowPopup(true); // Popup handles refresh + back
      } catch (error) {
ErrorService.logError(
  "NewsEditForm.onDelete",
  error
);        setPopupType("danger");
        setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
);
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  };

  // reassign
  const onReAssign = async () => {
    if (!editId) return;

    if (!form.Comments || !form.Comments.trim()) {
      setPopupType("danger");
      setPopupMessage("Comment is required before this action.");
      setShowPopup(true);
      return;
    }

    const creatorId = form.AuthorId;
    const approverId = assignedtoId;

    if (!creatorId || !approverId) {
      
      return;
    }

    // Show confirm popup
    setPopupType("confirm");
    setPopupMessage(
      `Request Number: ${formatRequestNumber(
        editId,
      )} - Are you sure you want to Re-Assign this Request?`,
    );

    setConfirmAction(() => async () => {
      try {
        // --- LM → Send back to Creator ---
        // if (isApprover && isPendingLM && isPendingWithMe) {
        if (canShowLMButtons) {
          await service.updateNews(editId, {
            Status: "PendingWithCreator",
            PendingWithId: creatorId,
            Comments: form.Comments,
          });
        }

        // --- CM → Send back to LM ---
        else if (isPrimaryCM && isPendingCM && isPendingWithMe) {
          await service.updateNews(editId, {
            Status: "PendingWithLM",
            PendingWithId: approverId,
            Comments: form.Comments,
            IsReassigned: true,
          });
        }

        // 🔹 IMPORTANT — reset confirm state first
        setConfirmAction(null);

        // 🔹 Show success popup (this was missing before)
        setPopupType("success");
        setPopupMessage(
          `Request Number: ${formatRequestNumber(
            editId,
          )} - Re-Assigned Successfully.`,
        );
        setShowPopup(true);
      } catch (error) {
ErrorService.logError(
  "NewsEditForm.onReAssign",
  error
);        setConfirmAction(null);
        setPopupType("danger");
       setPopupMessage(
  `Request Number: ${formatRequestNumber(editId)} - ${ErrorService.getUserMessage(error)}`
); 
        setShowPopup(true);
      }
    });

    setShowPopup(true);
  };

  const getExpiryPreview = () => {
    const dt = new Date();
    dt.setDate(dt.getDate() + 90);

    const yyyy = dt.getFullYear();
    const mm = ("0" + (dt.getMonth() + 1)).slice(-2);
    const dd = ("0" + dt.getDate()).slice(-2);

    return `${yyyy}-${mm}-${dd}`;
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

  // Allowed characters (A–Z a–z 0–9 + approved specials)
  //const allowedRegex = /^[A-Za-z0-9()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;
  const allowedRegex = /^[A-Za-z0-9\u0900-\u097F()&_\-+=\[\]{}:;'",.?/₹$!@%*\s]*$/;


  // Helper for input change
  const handleValidatedInput = (
    value: string,
    updateFn: (val: string) => void,
  ) => {
    if (!allowedRegex.test(value)) return;
    updateFn(value);
  }; 

  // Helper for paste protection
  const handleValidatedPaste = (
    e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>,
  ) => {
    const paste = e.clipboardData.getData("text");
    if (!allowedRegex.test(paste)) {
      e.preventDefault();
    }
  };

  return (
    <section className={`${styles.editNews} py-4`}>
      <Card className={`${styles.cardCommon}  p-1 p-sm-3 p-md-4`}>
        <Card.Body>
          <h4 className={`${styles.pageTitle}`}>{"Edit News"}</h4>

          <Form>
            <Row className="g-3">
              {/* Title */}
              <Col xs={12}>
                <Form.Group controlId="title" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Title <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Enter title"
                    value={form.Title || ""}
                    maxLength={255}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Title: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                    required
                  />
                  <small style={{ color: (form.Title || "").length > 250 ? "red" : "#666" }}>
  {(form.Title || "").length}/255 characters
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
                      value={form?.MainDescription || ""}
                      readOnly={!canEditFields}
                      onChange={(val: string) => {
                        if (!canEditFields) return;
                        setForm({ ...form, MainDescription: val });
                      }}
                      modules={modules}
                      formats={formats}
                      placeholder="Enter news content"
                    />
                  </div>
                </Form.Group>
              </Col>

              {/* Publish In */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="publishIn" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Publish In <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    required
                    value={form.PublishIn || ""}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      setForm({ ...form, PublishIn: e.target.value });
                    }}
                  >
                    <option value="">Select</option>
                    <option value="corporate">Corporate</option>
                    <option value="sbu">My Team</option>
                  </Form.Select>
                </Form.Group>
              </Col>

              {/* News Type */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="newsType" className="">
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
                    initialValues={newsType}
                    onChange={(terms) => {
                      setNewsType(terms || []);

                      if (terms && terms.length > 0) {
                        const selected = terms[0];

                        setForm({
                          ...form,
                          NewsTypes: {
                            Label: selected.name,
                            TermGuid: selected.key,
                          },
                        });
                      } else {
                        setForm({
                          ...form,
                          NewsTypes: null,
                        });
                      }
                    }}
                    isTermSetSelectable={false}
                    disabled={!canEditFields}
                  />
                </Form.Group>
              </Col>

              {/* ConfidentialityIndex */}

              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Confidentiality Index
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={form.ConfidentialityIndex?.Label || "Confidential"}
                    readOnly
                  />
                </Form.Group>
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              {/* ArchivalPolicy */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Archival Policy
                  </Form.Label>

                  <Form.Control
                    disabled
                    className={styles.inputBox}
                    type="text"
                    value={
                      form.ArchivalPolicy?.Label || "As Per Company Policy"
                    }
                    readOnly
                  />
                </Form.Group>
              </Col>

              {/* EventDate */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="eventDate" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Event Date <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="date"
                    value={form.EventDate || ""}
                    disabled={!canEditFields}
                    onChange={(e) => {
                      setForm({ ...form, EventDate: e.target.value });
                    }}
                    required
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
                    value={form.ExpiryDate || getExpiryPreview()}
                    readOnly
                  />
                </Form.Group>
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              {/* Approver */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Approver <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    value={form.AssignTo?.[0] || ""}
                    disabled={disableWorkflowDropdowns}
                    onChange={(e) => {
                      const val = e.target.value;

                      setForm({
                        ...form,
                        AssignTo: val ? [Number(val)] : [],
                      });
                    }}
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

              {/* Content Manager */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group>
                  <Form.Label className={styles.floatingLabelStyle}>
                    Content Manager <span className={styles.required}>*</span>
                  </Form.Label>

                  <Form.Select
                    className={styles.inputBox}
                    required
                    value={selectedContentManager}
                    disabled={disableWorkflowDropdowns}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedContentManager(val);

                      setForm({
                        ...form,
                        FollowOnActionBy: val ? [Number(val)] : [],
                      });
                    }}
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

                {canEditFields ? (
                  <PeoplePickerCustom
                    context={context}
                    isClearable={true}
                    selectedValue={viewBy}
                    onChange={(val) => {
                      setViewBy(val);
                      setForm({
                        ...form,
                        ViewBy: val ? [Number(val.value)] : [],
                      });
                    }}
                  />
                ) : (
                  <Form.Group>
                    <Form.Control
                      className={styles.inputBox}
                      type="text"
                      disabled
                      value={viewBy?.label || ""}
                    />
                  </Form.Group>
                )}
              </Col>
            </Row>

            {/* InformToBrand */}
            <Row className="mt-2">
              <Col>
                <Form.Check
                  inline
                  type="checkbox"
                  id="informToBrand"
                  label="Inform To Brand (would you like brand team to take your publications)"
                  checked={!!form.InformToBrand}
                  disabled={!canEditFields}
                  onChange={(e) => {
                    setForm({ ...form, InformToBrand: e.target.checked });
                  }}
                />
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              {/* Thumbnail */}
              <Col xs={12} md={6}>
                <Form.Group controlId="thumbnail">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Thumbnail <span className={styles.required}>*</span>{" "}
                    <small>
                      (Image should be less than 2MB and allowed bmp, gif, png,
                      jpg, jpeg)
                    </small>
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    disabled={!canEditFields}
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                  onChange={async (
  e: React.ChangeEvent<HTMLInputElement>,
) => { 

  if (!canEditFields) return;

  const file = e.target.files?.[0];

  if (!file) return;

  const isValid = await validateImageFile(
    file,
    e.target
  );

  if (!isValid) return;

  setThumbnailFile(file);

  setForm({ 
    ...form,
    Thumbnail: file.name
  });
}}
                  />

                  {form.Thumbnail && (
                    <div className="mt-2">
                      <small className="text-muted">
                        Saved File: {form.Thumbnail}
                      </small>
                    </div>
                  )}
                </Form.Group>
              </Col>

              {/* ThumbnailCaption */}
              <Col xs={12} md={6}>
                <Form.Group controlId="thumbnailCaption">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Thumbnail Caption
                  </Form.Label>
                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    placeholder="Thumbnail caption"
                    value={form.ThumbnailCaption || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, ThumbnailCaption: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Pictures Upload */}
            <Row className="mt-1 g-3">
              {/* Picture1 */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture1">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 1
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    disabled={!canEditFields}
                    onChange={async (e: React.ChangeEvent<HTMLInputElement>) => {
                      if (!canEditFields) return;

                      const file = e.target.files?.[0];
                      if (!file) return;

                      const isValid = await validateImageFile(file, e.target);
                      if (!isValid) return;

                      setPic1File(file);
                      setForm({ ...form, Picture1: file.name });
                    }}
                  />

                  {form.Picture1 && (
                    <div className="mt-2">
                      <small className="text-muted">
                        Saved File: {form.Picture1}
                      </small>
                    </div>
                  )}
                </Form.Group>
              </Col>

              {/* Picture2 */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture2">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 2
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    disabled={!canEditFields}
                    onChange={async (e: React.ChangeEvent<HTMLInputElement>) => {
                      if (!canEditFields) return;

                      const file = e.target.files?.[0];
                      if (!file) return;

                      const isValid = await validateImageFile(file, e.target);
                      if (!isValid) return;

                      setPic2File(file);
                      setForm({ ...form, Picture2: file.name });
                    }}
                  />

                  {form.Picture2 && (
                    <div className="mt-2">
                      <small className="text-muted">
                        Saved File: {form.Picture2}
                      </small>
                    </div>
                  )}
                </Form.Group>
              </Col>

              {/* Picture3 */}
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="picture3">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Upload Picture 3
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="file"
                    accept=".bmp,.gif,.png,.jpg,.jpeg"
                    disabled={!canEditFields}
                    onChange={async (e: React.ChangeEvent<HTMLInputElement>) => {
                      if (!canEditFields) return;

                      const file = e.target.files?.[0];
                      if (!file) return;

                      const isValid = await validateImageFile(file, e.target);
                      if (!isValid) return;

                      setPic3File(file);
                      setForm({ ...form, Picture3: file.name });
                    }}
                  />

                  {form.Picture3 && (
                    <div className="mt-2">
                      <small className="text-muted">
                        Saved File: {form.Picture3}
                      </small>
                    </div>
                  )}
                </Form.Group>
              </Col>
            </Row>

            {/* Captions */}
            <Row className="mt-1 g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pic1Caption" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 1 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.Pic1Caption || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Pic1Caption: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pic2Caption" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 2 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.Pic2Caption || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Pic2Caption: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="pic3Caption" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Picture 3 Caption
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.Pic3Caption || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Pic3Caption: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Additional info + Comments */}

            <Row className="mt-1 g-3">
              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="additionalInfo1" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Additional Info 1
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.AdditionalInformation1 || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, AdditionalInformation1: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="additionalInfo2" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Additional Info 2
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.AdditionalInformation2 || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, AdditionalInformation2: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>

              <Col xs={12} sm={6} md={4}>
                <Form.Group controlId="additionalInfo3" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Additional Info 3
                  </Form.Label>

                  <Form.Control
                    className={styles.inputBox}
                    type="text"
                    value={form.AdditionalInformation3 || ""}
                    disabled={!canEditFields}
                    onChange={(e) =>
                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, AdditionalInformation3: val }),
                      )
                    }
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>
            </Row>

            <Row className="mt-1 g-3">
              <Col>
                <Form.Group controlId="comments" className="">
                  <Form.Label className={styles.floatingLabelStyle}>
                    Comment
                  </Form.Label>

                  <Form.Control
                    as="textarea"
                    ref={commentRef}
                    value={form.Comments || ""}
                    disabled={!canEditComment}
                    style={{
                      minHeight: "50px",
                      resize: "none",
                      overflow: "hidden",
                    }}
                    onChange={(e) => {
                      if (!canEditComment) return;

                      e.target.style.height = "auto";
                      e.target.style.height = `${Math.max(e.target.scrollHeight, 50)}px`;

                      handleValidatedInput(e.target.value, (val) =>
                        setForm({ ...form, Comments: val }),
                      );
                    }}
                    onPaste={handleValidatedPaste}
                  />
                </Form.Group>
              </Col>
            </Row>

            {/* Navigation Links */}
            <Row className="mt-4">
              {/* {isStaffPosting && transferListUrl && imagesFolderUrl && ( */}
              {isStaffPosting && imagesFolderUrl && (
                <div className="mt-3 p-3 border rounded">
                  <h6 className="fw-bold mb-2">Navigation Links</h6>
                  <div>
                    TransferListImages Folder :
                    <span
                      className="ms-2 text-primary"
                      style={{ cursor: "pointer", textDecoration: "underline" }}
                      onClick={() => {
                        if (editId) {
                          openImagesLibraryFolder(editId);
                        }
                      }}
                    >
                      Open Folder
                    </span>
                  </div>
                  <div className="mb-2">
                    TransferSbuIndexList :
                    <a
                      href="#"
                      className="ms-2"
                      target="_blank"
                      onClick={(e) => {
                        e.preventDefault();
                        if (editId) {
                          openTransferFolder(editId); // this opens new tab
                        }
                      }}
                    >
                      Open List
                    </a>
                  </div>
                </div> 
              )}
            </Row>

            {/* Buttons */}
            <Row className="mt-4">
              <div className="d-flex flex-column flex-md-row gap-2">
                <Button
                  className={`${styles.primaryOutlineBtn} btn`}
                  onClick={onPreviewPopup}
                >
                  Preview
                </Button>

                {canShowSaveButton && (
                  <Button
                    className={`${styles.primaryBtn} btn`}
                    onClick={onSaveDraft}
                  >
                    Save
                  </Button>
                )}

                <div className="ms-md-auto d-flex flex-column flex-md-row gap-2">
                  {canShowApproveRejectButtons && (
                    <>
                      {canShowReassignButton && (
                        <Button
                          className={styles.reassignBtn}
                          onClick={onReAssign}
                        >
                          ReAssign
                        </Button>
                      )}

                      <Button className={styles.approveBtn} onClick={onApprove}>
                        Approve
                      </Button>

                      <Button className={styles.rejectBtn} onClick={onReject}>
                        Reject
                      </Button>
                    </>
                  )}

                  {canShowDeleteButton && (
                    <Button
                      // variant="outline-danger"
                      className={styles.deleteBtn}
                      onClick={onDelete}
                    >
                      Delete
                    </Button>
                  )}

                  <Button
                    className={`${styles.primaryOutlineBtn} btn`}
                    onClick={onBack}
                  >
                    Cancel
                  </Button>

                  {canShowSubmitButton && (
                    <Button
                      className={`${styles.primaryBtn} btn`}
                      onClick={onSubmit}
                    >
                      Submit
                    </Button>
                  )}
                </div>
              </div>
            </Row>
          </Form>
        </Card.Body>
      </Card>

      {/* Popup Modal */}
      <PopupModal
        show={showPopup}
        type={popupType}
        message={popupMessage}
        onClose={async (result?: "yes" | "no") => {
          // Cancel clicked
          if (popupType === "confirm" && result === "no") {
            setConfirmAction(null);
            setShowPopup(false);
            return;
          }

          // Confirm YES clicked
          if (popupType === "confirm" && result === "yes" && confirmAction) {
            const action = confirmAction;
            setConfirmAction(null); //  clear first
            setShowPopup(false); //  close confirm popup FIRST
            await action(); //  run reassignment → will show SUCCESS popup
            return;
          }

          // Normal close (OK button)
          setShowPopup(false);

          if (popupType === "success") {
            if (refreshList) await refreshList();
            onBack();
          }
        }}
      />
    </section>
  );
};

export default NewsEditForm;
