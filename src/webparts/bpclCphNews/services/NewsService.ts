import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import "@pnp/sp/attachments";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/profiles";

import { SPHttpClient } from "@microsoft/sp-http";

import AuditService from "../services/AuditService";
import ErrorService from "./ErrorService";

export class NewsService {
  public sp;
  private context: any;

  private auditService: AuditService;

  constructor(context: any) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
    this.auditService = new AuditService(
      this.sp,
      context.pageContext.web.absoluteUrl,
    );
  }
  // getting sita collection admin
  public async isCurrentUserSiteAdmin(): Promise<boolean> {
    try {
      const user = await this.sp.web.currentUser();

      // IsSiteAdmin = true only for Site Collection Admin
      return user?.IsSiteAdmin === true;
    } catch (error) {
      ErrorService.logError("NewsService.isCurrentUserSiteAdmin", error);
      return false;
    }
  }
  // getting dashboard card count
  public async getAllCounts(): Promise<{
    news: number;
    events: number;
    broadcasts: number;
    brands: number;
  }> {
    try {
      let news = 0;
      let events = 0;
      let broadcasts = 0;
      let brands = 0;

      const { profile, sbu } = await this.getUserContext();

      let filter = "";

       if (sbu) {
        let teamName = (sbu || "").trim();

        if (
          teamName === "HRS" ||
          teamName === "HRD" ||
          teamName === "HUMAN RESOURCES"
        ) {
          filter = "(SBU eq 'HRS' or SBU eq 'HRD' or SBU eq 'HUMAN RESOURCES')";
        } else {
          filter = `SBU eq '${teamName.replace(/'/g, "''")}'`;
        }
      } else if (profile.department) {
        filter = `Department eq '${profile.department.replace(/'/g, "''")}'`;
      }

      if (sbu) {
  const teamName = sbu.trim();
  filter = `SBU eq '${teamName.replace(/'/g, "''")}'`;
} else if (profile.department) {
  filter = `Department eq '${profile.department.replace(/'/g, "''")}'`;
}

      if (!filter) {
        return { news: 0, events: 0, broadcasts: 0, brands: 0 };
      }

      const items = this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.select("CommunicationType")
        .filter(`${filter} and Status eq 'Published'`)
        .top(5000);

      //  PnP v3 paging
      for await (const batch of items) {
        batch.forEach((item: any) => {
          const type = item.CommunicationType;

          if (type === "News") news++;
          else if (type === "Event") events++;
          else if (type === "BroadCast") broadcasts++;
          else if (type === "Brand") brands++;
        });
      }

      return { news, events, broadcasts, brands };
    } catch (error) {
      ErrorService.logError("NewsService.getAllCounts", error);

      return { news: 0, events: 0, broadcasts: 0, brands: 0 };
    }
  }

  public async getLoggedInUserProfileInfo(): Promise<{
    department: string;
    reportingManagerLogin: string;
    reportingManagerUser?: any;
  }> {
    try {
      const profile = await this.sp.profiles.myProperties();

      const dept =
        profile.UserProfileProperties?.find(
          (p: any) => p.Key === "Department" || p.Key === "SPS-Department",
        )?.Value || "";

      const managerLogin =
        profile.UserProfileProperties?.find(
          (p: any) => p.Key === "Manager" || p.Key === "SPS-Manager",
        )?.Value || "";

      let managerUser = null;

      if (managerLogin) {
        try {
          managerUser = await this.sp.web.ensureUser(managerLogin);
        } catch (error) {
          ErrorService.logError("NewsService.ensureUser", error);
        }
      }

      return {
        department: dept,
        reportingManagerLogin: managerLogin,
        reportingManagerUser: managerUser,
      };
    } catch (error) {
      ErrorService.logError("NewsService.getLoggedInUserProfileInfo", error);

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  private async getUserGroups(): Promise<Set<string>> {
    const client = await this.context.msGraphClientFactory.getClient("3");

    let requestUrl: string | null =
      "/me/transitiveMemberOf?$select=displayName";

    const userGroups = new Set<string>();

    while (requestUrl) {
      const graphResponse: any = await client.api(requestUrl).get();

      if (graphResponse.value) {
        graphResponse.value.forEach((g: { displayName?: string }) => {
          if (g.displayName) {
            userGroups.add(g.displayName.trim().toLowerCase());
          }
        });
      }

      requestUrl = graphResponse["@odata.nextLink"]
        ? graphResponse["@odata.nextLink"].replace(
            "https://graph.microsoft.com/v1.0",
            "",
          )
        : null;
    }

    //     alert(
    //   "Groups = " +
    //   Array.from(userGroups).join(", ")
    // );

    return userGroups;
  }

  // getting sbu values from ad
  public async getCurrentUserSBUFromGraph(): Promise<string> {
    //    const isTestMode = true; // change to false after testing

    // if (isTestMode) {
    //   return "MUMBAI REFINERY";
    // } 

    try {
      const client = await this.context.msGraphClientFactory.getClient("3");

      const user = await client
        .api("/me?$select=onPremisesExtensionAttributes")
        .get();

      return (
        user?.onPremisesExtensionAttributes?.extensionAttribute4?.trim() || ""
      );
    } catch (error) {
      ErrorService.logError("NewsService.getCurrentUserSBUFromGraph", error);

      return "";
    }
  } 

  // private buildTeamFilter(profile: any, sbu: string): string {
  //   let teamName = (sbu || "").trim();
  //   let department = (profile?.department || "").trim();

  //   // Normalize HR
  //   if (
  //     teamName === "HRS" ||
  //     teamName === "HRD" 
  //     // teamName === "HUMAN RESOURCES"
  //   ) {
  //     teamName = "HUMAN RESOURCES";
  //   }

  //   const safeTeam = teamName.replace(/'/g, "''");
  //   const safeDept = department.replace(/'/g, "''");

  //   if (safeTeam && safeDept) {
  //     return `(TeamName/TeamName eq '${safeTeam}' or TeamName/TeamName eq '${safeDept}')`;
  //   } else if (safeTeam) {
  //     return `TeamName/TeamName eq '${safeTeam}'`;
  //   } else if (safeDept) {
  //     return `TeamName/TeamName eq '${safeDept}'`;
  //   }

  //   return "";
  // }
  
private buildTeamFilter(profile: any, sbu: string): string {
  let teamName = (sbu || "").trim();
    let department = (profile?.department || "").trim();

  const isHR = ["HRS", "HRD", "HUMAN RESOURCES"].includes(teamName);

  const safeDept = department.replace(/'/g, "''");

  const teamFilter = isHR
    ? "TeamName/TeamName eq 'HUMAN RESOURCES'"
    : teamName
      ? `TeamName/TeamName eq '${teamName.replace(/'/g, "''")}'`
      : "";

  const departmentFilter = safeDept
    ? `TeamName/TeamName eq '${safeDept}'`
    : "";

  if (teamFilter && departmentFilter) {
    return `(${teamFilter} or ${departmentFilter})`;
  }

  return teamFilter || departmentFilter;
}

  // private isSameUser(u: any, currentUser: any): boolean {
  //   return (
  //     u.Id === currentUser.Id ||
  //     u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase()
  //   );
  // } 

  
private isSameUser(u: any, currentUser: any): boolean {
  if (!u || !currentUser) {
    return false;
  }

  if (
    u.Id != null &&
    currentUser.Id != null &&
    Number(u.Id) === Number(currentUser.Id)
  ) {
    return true;
  }

  const assignedEmail = (u.EMail || u.Email || "")
    .trim()
    .toLowerCase();

  const currentEmail = (currentUser.Email || currentUser.EMail || "")
    .trim()
    .toLowerCase();

  return !!assignedEmail &&
    !!currentEmail &&
    assignedEmail === currentEmail;
}


  private applyTeamPriority(
    items: any[],
    teamName: string,
    department: string,
    currentUser: any,
  ) {
    if (!items || items.length === 0) return [];

    const normalize = (v: string) => (v || "").trim().toLowerCase();

    // Normalize HR values
    if (
      normalize(teamName) === "hrs" ||
      normalize(teamName) === "hrd" ||
      normalize(teamName) === "human resources"
    ) {
      teamName = "HUMAN RESOURCES";
    }

    // const isSameUser = (u: any) =>
    //   u.Id === currentUser.Id ||
    //   u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

    // const isUserInRow = (row: any) => {
    //   const allUsers = [
    //     ...(row.PrimaryContentManager || []),
    //     ...(row.Approvers || []),
    //     ...(row.NewsCreators || []),
    //     ...(row.BroadcastCreators || []),
    //     ...(row.EventCreators || []),
    //     ...(row.BrandCreators || []),
    //   ];
    //   return allUsers.some(isSameUser);
    // };

    // 🔹 STEP 1: Try SBU match
    let matchedItems = [];

    if (teamName) {
      matchedItems = items.filter(
        (i) => normalize(i.TeamName?.TeamName) === normalize(teamName),
      );
    }

    // 🔹 STEP 2: If NO SBU match → fallback to Department
    if (matchedItems.length === 0 && department) {
      matchedItems = items.filter(
        (i) => normalize(i.TeamName?.TeamName) === normalize(department),
      );
    }

    // 🔹 STEP 3: Filter by user
    //return matchedItems.filter(isUserInRow);
    return matchedItems;
  }

  // private async getUserContext() {
  //   if (this._userContext) return this._userContext;

  //   const [currentUser, profile, sbu] = await Promise.all([
  //     this.sp.web.currentUser(),
  //     this.getLoggedInUserProfileInfo(),
  //     this.getCurrentUserSBUFromGraph(),
  //   ]);

  //   this._userContext = {
  //     currentUser,
  //     profile,
  //     sbu,
  //   };

  //   return this._userContext;
  // }

// private isTestMode = true; // Set false after testing

// private testUserEmail = "amankumar002@bharatpetroleum.in";
// //private testUserTitle = "Aman Kumar";
// private testDept = "RETAIL TERR.OPS.BHATINDA DEPOT";
// private testSbu = "RETAIL"; // Change to "IS-Dummy" to test IS

private _userContext: any = null;
private async getUserContext() {
  if (this._userContext) return this._userContext;

  const [currentUser, profile, sbu] = await Promise.all([
    this.sp.web.currentUser(),
    this.getLoggedInUserProfileInfo(),
    this.getCurrentUserSBUFromGraph(),
  ]);

  this._userContext = {
    currentUser,
    profile,
    sbu,
  };


  return this._userContext;
}


// private async getUserContext() {
//   if (this._userContext) return this._userContext;

//   const [currentUser, profile, actualSbu] = await Promise.all([
//     this.sp.web.currentUser(),
//     this.getLoggedInUserProfileInfo(),
//     this.getCurrentUserSBUFromGraph(),
//   ]);

//   // Keep actual identity unchanged.
//   // Override only SBU and Department for testing.
//   const effectiveSbu = this.isTestMode
//     ? this.testSbu
//     : actualSbu;

//   const effectiveProfile = {
//     ...profile,
//     department: this.isTestMode
//       ? this.testDept
//       : profile.department,
//   };

//   this._userContext = {
//     currentUser,
//     profile: effectiveProfile,
//     sbu: effectiveSbu,
//   };

//   console.log("========== NEWS TEST MODE ==========");
//   console.log("Test mode enabled:", this.isTestMode);
//   console.log("Actual logged-in email:", currentUser.Email);
//   console.log("Actual SharePoint user ID:", currentUser.Id);
//   console.log("Test target email (reference only):", this.testUserEmail);
//   console.log("Actual SBU from Graph:", actualSbu);
//   console.log("Effective SBU:", effectiveSbu);
//   console.log("Actual department:", profile.department);
//   console.log("Effective department:", effectiveProfile.department);
//   console.log("====================================");

//   return this._userContext;
// }

  public async isUserPartOfSbuTeam(): Promise<boolean> {
    const userGroups = await this.getUserGroups();

    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);
    console.log("filterQuery:", sbu,
      profile.department, filterQuery
    );

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "TeamName/TeamName",
        "NewsCreators/Id",
        // "NewsCreators/EMail",
        "NewsCreators/Title",

        "Approvers/Id",
          // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand("TeamName", "NewsCreators", "Approvers", "PrimaryContentManager")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );
  
  //   return finalItems.some((row: any) => {
  //     const match = (users: any[]) =>
  //       users?.some((u: any) => {
  //         const isDirectUser =
  //           u.Id === currentUser.Id ||
  //           u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

  //         const isGroupMember =
  //           u.Title && userGroups.has(u.Title.trim().toLowerCase());
  // console.log("Checking user:", u.Id+u.Title+u.EMail, "Current User:", currentUser.Id);
  //         return isDirectUser || isGroupMember;
  //       }); 

  //     return (
  //       match(row.NewsCreators) ||
  //       match(row.Approvers) ||
  //       match(row.PrimaryContentManager)
  //     );
  //   });
  
return finalItems.some((row: any) => {
  const match = (users: any[] | undefined) =>
    this.normalizeUsers(users).some((u: any) => {
      const isDirectUser = this.isSameUser(u, currentUser);

      const isGroupMember =
        !!u.Title &&
        userGroups.has(u.Title.trim().toLowerCase());

      console.log("[User Identity Check]", {
        listUserId: u.Id,
        listUserTitle: u.Title,
        listUserEmail: u.EMail,
        currentUserId: currentUser.Id,
        currentUserEmail: currentUser.Email,
        idMatches:
          Number(u.Id) === Number(currentUser.Id),
        emailMatches:
          !!u.EMail &&
          u.EMail.trim().toLowerCase() ===
            (currentUser.Email || "").trim().toLowerCase(),
        isDirectUser,
        isGroupMember,
        rowTeamName: row.TeamName?.Title || row.TeamName,
      });

      return isDirectUser || isGroupMember;
    });

  return (
    match(row.NewsCreators) ||
    match(row.Approvers) ||
    match(row.PrimaryContentManager)
  );
});

  }

  private async validateFileBeforeUpload(file: File): Promise<boolean> {
    const allowedExtensions = ["bmp", "gif", "png", "jpg", "jpeg"];

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
      return false;
    }

    // MIME type validation
    if (!allowedMimeTypes.includes(file.type)) {
      return false;
    }

    // File size validation (2MB)
    if (file.size > 2 * 1024 * 1024) {
      return false;
    }

    // Magic byte validation
    const buffer = await file.slice(0, 8).arrayBuffer();

    const bytes = new Uint8Array(buffer);

    const isPng =
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47;

    const isJpg = bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;

    const isGif = bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46;

    const isBmp = bytes[0] === 0x42 && bytes[1] === 0x4d;

    return isPng || isJpg || isGif || isBmp;
  }

  // getting current user role
  private async getCurrentUserRoles(): Promise<string[]> {
    const userGroups = await this.getUserGroups();
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "TeamName/TeamName",
        "NewsCreators/Id",
        // "NewsCreators/EMail",
        "NewsCreators/Title",

        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand("TeamName", "NewsCreators", "Approvers", "PrimaryContentManager")
      .filter(filterQuery)();

    const roles: string[] = [];

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    // const currentUserId = currentUser.Id;

    // const isSameUser = (u: any) =>
    //   u.Id === currentUserId ||
    //   u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

    // for (const item of finalItems) {
      // if (item.NewsCreators?.some((u: any) => isSameUser(u))) {
      //   roles.push("Creator");
      // }
    //   for (const u of item.NewsCreators || []) {
    //     const isDirectUser =
    //       u.Id === currentUserId ||
    //       u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

    //     const isGroupMember =
    //       u.Title && userGroups.has(u.Title.trim().toLowerCase());

    //     if (isDirectUser || isGroupMember) {
    //       roles.push("Creator");
    //       break;
    //     }
    //   }

    //   if (item.Approvers?.some((u: any) => isSameUser(u))) {
    //     roles.push("Approver");
    //   }

    //   if (item.PrimaryContentManager?.some((u: any) => isSameUser(u))) {
    //     roles.push("Content Manager");
    //   }
    // }
    
for (const item of finalItems) {
  const creators = this.normalizeUsers(item.NewsCreators);

  const isCreator =
    creators.some((u: any) => this.isSameUser(u, currentUser)) ||
    creators.some(
      (u: any) =>
        !!u.Title &&
        userGroups.has(u.Title.trim().toLowerCase())
    );

  if (isCreator) {
    roles.push("Creator");
  }

  if (
    this.normalizeUsers(item.Approvers).some(
      (u: any) => this.isSameUser(u, currentUser)
    )
  ) {
    roles.push("Approver");
  }

  if (
    this.normalizeUsers(item.PrimaryContentManager).some(
      (u: any) => this.isSameUser(u, currentUser)
    )
  ) {
    roles.push("Content Manager");
  }
}

    // ES5-safe unique logic
    const uniqueRoles: string[] = [];

    roles.forEach((role) => {
      if (uniqueRoles.indexOf(role) === -1) {
        uniqueRoles.push(role);
      }
    });

    return uniqueRoles.length ? uniqueRoles : ["User"];
  }

  //  Create news in to the list
  public async saveNews(
    title: string,
    description: string,
    publishIn: string,
    eventDate: string,
    addInfo1: string,
    addInfo2: string,
    addInfo3: string,
    comments: string,
    thumbnailFile: File | null,
    thumbnailCaption: string,
    pic1: File | null,
    pic2: File | null,
    pic3: File | null,
    picCap1: string,
    picCap2: string,
    picCap3: string,
    approverId: string,
    contentManagerId: string,
    newsType: any,
    viewById: number | null,
    informToBrand: boolean,
    isTransferList: boolean,
    department: string,
    confidentialityIndex: any,
    archivalPolicy: any,
  ): Promise<{ itemId: number; requestNumber: string }> {
    try {
      const sbuText = await this.getCurrentUserSBUFromGraph();
      const result = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.add({
          Title: title,
          MainDescription: description,
          PublishIn: publishIn,
          EventDate: eventDate ? new Date(eventDate) : null,
          AdditionalInformation1: addInfo1,
          AdditionalInformation2: addInfo2,
          AdditionalInformation3: addInfo3,
          Comments: comments,
          Thumbnail: thumbnailFile ? thumbnailFile.name : "",
          ThumbnailCaption: thumbnailCaption,
          Picture1: pic1?.name || "",
          Picture2: pic2?.name || "",
          Picture3: pic3?.name || "",
          Pic1Caption: picCap1,
          Pic2Caption: picCap2,
          Pic3Caption: picCap3,
          NewsTypes: newsType
            ? {
                Label: newsType.name,
                TermGuid: newsType.key,
                WssId: -1,
              }
            : null,

          ConfidentialityIndex: confidentialityIndex
            ? this.getTaxonomyValue(confidentialityIndex)
            : null,

          ArchivalPolicy: archivalPolicy
            ? this.getTaxonomyValue(archivalPolicy)
            : null,

          AssignToId: approverId ? Number(approverId) : null,
          PendingWithId: approverId ? Number(approverId) : null,
          FollowOnActionById: contentManagerId
            ? Number(contentManagerId)
            : null,
          ViewById: viewById ? Number(viewById) : null,
          CommunicationType: "News",
          SBU: sbuText,
          InformToBrand: informToBrand,
          IsTransferList: isTransferList,
          Department: department,
          PublishedDate: new Date().toISOString(),
        });

      // Correct way to get ItemId
      const itemId = result.Id;
      if (!itemId) {
        throw new Error("Item created but ID could not be resolved.");
      }

      //  Generate RequestNumber
      const strId = itemId.toString();
      const padded = ("00000000" + strId).slice(-8); // 8 digits
      const requestNumber = `RQ${padded}`;

      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(itemId)
        .update({
          RequestNumber: requestNumber,
        });

      // Upload attachments

      const baseUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('CorpCommunication')/items(${itemId})/AttachmentFiles/add(FileName='`;

      // if (thumbnailFile) {
      //   await this.context.spHttpClient.post(
      //     `${baseUrl}${encodeURIComponent(thumbnailFile.name)}')`,
      //     SPHttpClient.configurations.v1,
      //     { body: thumbnailFile },
      //   );
      // }

      if (thumbnailFile) {
        const valid = await this.validateFileBeforeUpload(thumbnailFile);

        if (!valid) {
          throw new Error("Invalid image file detected.");
        }

        await this.context.spHttpClient.post(
          `${baseUrl}${encodeURIComponent(thumbnailFile.name)}')`,
          SPHttpClient.configurations.v1,
          { body: thumbnailFile },
        );
      }

      if (pic1) {
        const valid = await this.validateFileBeforeUpload(pic1);

        if (!valid) {
          throw new Error("Invalid image file detected.");
        }

        await this.context.spHttpClient.post(
          `${baseUrl}${encodeURIComponent(pic1.name)}')`,
          SPHttpClient.configurations.v1,
          { body: pic1 },
        );
      }

      if (pic2) {
        const valid = await this.validateFileBeforeUpload(pic2);

        if (!valid) {
          throw new Error("Invalid image file detected.");
        }

        await this.context.spHttpClient.post(
          `${baseUrl}${encodeURIComponent(pic2.name)}')`,
          SPHttpClient.configurations.v1,
          { body: pic2 },
        );
      }

      if (pic3) {
        const valid = await this.validateFileBeforeUpload(pic3);

        if (!valid) {
          throw new Error("Invalid image file detected.");
        }

        await this.context.spHttpClient.post(
          `${baseUrl}${encodeURIComponent(pic3.name)}')`,
          SPHttpClient.configurations.v1,
          { body: pic3 },
        );
      }

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      await this.auditService.log({
        action: "Create",
        operation: "POST",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: itemId,
        userRole: roleString,
      });

      return { itemId, requestNumber };
    } catch (error) {
      ErrorService.logError("NewsService.saveNews", error);

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      //  FAILURE LOG
      await this.auditService.log({
        action: "Create Failed",
        operation: "POST",
        statusCode: 500,
        listName: "CorpCommunication",
        userRole: roleString,
      });

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  public async getExpiryDateById(id: number): Promise<string | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .select("ExpiryDate")();

      return item?.ExpiryDate || null;
    } catch (error) {
      ErrorService.logError("NewsService.getExpiryDateById", error);

      return null;
    }
  }

  private normalizeUsers(field: any): any[] {
    if (!field) return [];
    return Array.isArray(field) ? field : [field];
  }

  public async getApproversByTeam(): Promise<any[]> {
    const { profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "Approvers/Id",
        "Approvers/Title",
      )
      .expand("TeamName", "Approvers")
      .filter(filterQuery)
      .top(500)();

    // const normalize = (v: string) => (v || "").trim().toLowerCase();

    // let finalItems = items.filter(
    //   (i: any) => normalize(i.TeamName?.TeamName) === normalize(sbu),
    // );

    // if (finalItems.length === 0) {
    //   finalItems = items.filter(
    //     (i: any) =>
    //       normalize(i.TeamName?.TeamName) === normalize(profile.department),
    //   );
    // }
    
const normalize = (value: string) => {
  const normalized = (value || "").trim().toUpperCase();

  return ["HRS", "HRD", "HUMAN RESOURCES"].includes(normalized)
    ? "HUMAN RESOURCES"
    : normalized;
};

let finalItems = items.filter(
  (item: any) =>
    normalize(item.TeamName?.TeamName) === normalize(sbu)
);

if (finalItems.length === 0) {
  finalItems = items.filter(
    (item: any) =>
      normalize(item.TeamName?.TeamName) ===
      normalize(profile.department)
  );
}


    const users: any[] = [];

    finalItems.forEach((i: any) => {
      if (i.Approvers) {
        if (Array.isArray(i.Approvers)) {
          i.Approvers.forEach((u: any) => users.push(u));
        } else {
          users.push(i.Approvers);
        }
      }
    });

    return users;
  }

  public async getContentManagersByTeam(): Promise<any[]> {
    const { profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "PrimaryContentManager/Id",
        "PrimaryContentManager/Title",
      )
      .expand("TeamName", "PrimaryContentManager")
      .filter(filterQuery)
      .top(500)();

    // const normalize = (v: string) => (v || "").trim().toLowerCase();

    // let finalItems = items.filter(
    //   (i: any) => normalize(i.TeamName?.TeamName) === normalize(sbu),
    // );

    // if (finalItems.length === 0) {
    //   finalItems = items.filter(
    //     (i: any) =>
    //       normalize(i.TeamName?.TeamName) === normalize(profile.department),
    //   );
    // }
    
const normalize = (value: string) => {
  const normalized = (value || "").trim().toUpperCase();

  return ["HRS", "HRD", "HUMAN RESOURCES"].includes(normalized)
    ? "HUMAN RESOURCES"
    : normalized;
};

let finalItems = items.filter(
  (item: any) =>
    normalize(item.TeamName?.TeamName) === normalize(sbu)
);

if (finalItems.length === 0) {
  finalItems = items.filter(
    (item: any) =>
      normalize(item.TeamName?.TeamName) ===
      normalize(profile.department)
  );
}


    const users: any[] = [];

    finalItems.forEach((i: any) => {
      if (i.PrimaryContentManager) {
        if (Array.isArray(i.PrimaryContentManager)) {
          i.PrimaryContentManager.forEach((u: any) => users.push(u));
        } else {
          users.push(i.PrimaryContentManager);
        }
      }
    });

    return users;
  } 

  // is role check creator or not
  public async isCurrentUserNewsCreator(): Promise<boolean> {
    //const userId = await this.getCurrentUserId();

    const userGroups = await this.getUserGroups();

    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "TeamName/TeamName",
        "NewsCreators/Id",
        // "NewsCreators/EMail",
        "NewsCreators/Title",

        "Approvers/Id",
        // "Approvers/EMail",

        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
      )
      .expand("TeamName", "NewsCreators", "Approvers", "PrimaryContentManager")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );
console.log("Current user:", {
  id: currentUser.Id,
  email: currentUser.Email
});

console.log(
  "TeamCMList rows:",
  finalItems.map((row: any) => ({
    team: row.TeamName?.TeamName,
    creators: row.NewsCreators,
    approvers: row.Approvers,
    contentManagers: row.PrimaryContentManager
  }))
);
    // return finalItems.some((row: any) => {
    //   // const isCreator = row.NewsCreators?.some(
    //   //   (u: any) =>
    //   //     u.Id === userId ||
    //   //     u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
    //   // );

    //   let isCreator = false;
    //   for (const u of row.NewsCreators || []) {
    //     const isDirectUser =
    //       u.Id === userId ||
    //       u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase();

    //     const isGroupMember =
    //       u.Title && userGroups.has(u.Title.trim().toLowerCase());

    //     if (isDirectUser || isGroupMember) {
    //       isCreator = true;
    //       break;
    //     }
    //   }

    //   const isPrimaryCM = row.PrimaryContentManager?.some(
    //     (u: any) =>
    //       u.Id === userId ||
    //       u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
    //   );

    //   const isApprover = row.Approvers?.some(
    //     (u: any) =>
    //       u.Id === userId ||
    //       u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase(),
    //   );

    //   return isCreator || isPrimaryCM || isApprover;
    // });
    
return finalItems.some((row: any) => {
  const creators = this.normalizeUsers(row.NewsCreators);

  const isCreator =
    creators.some((u: any) => this.isSameUser(u, currentUser)) ||
    creators.some(
      (u: any) =>
        !!u.Title &&
        userGroups.has(u.Title.trim().toLowerCase())
    );

  const isPrimaryCM = this.normalizeUsers(
    row.PrimaryContentManager
  ).some((u: any) => this.isSameUser(u, currentUser));

  const isApprover = this.normalizeUsers(
    row.Approvers
  ).some((u: any) => this.isSameUser(u, currentUser));

  return isCreator || isPrimaryCM || isApprover;
});

  }

  // create button visiblity
  public async canUserCreateByType(
    type: "News" | "BroadCast" | "Event" | "Brand",
  ): Promise<boolean> {
    const userGroups = await this.getUserGroups();

    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);

 const items = await this.sp.web.lists
  .getByTitle("TeamCMList")
  .items.select(
    "Id",
    "IsNewsCreator",
    "IsBroadcastCreator",
    "IsEventCreator",
    "IsBrandCreator",

    "NewsCreators/Id",
    "NewsCreators/Title",

    "BroadcastCreators/Id",
    "BroadcastCreators/Title",

    "EventCreators/Id",
    "EventCreators/Title",

    "BrandCreators/Id",
    "BrandCreators/Title",

    "Approvers/Id",
    "Approvers/Title",

    "PrimaryContentManager/Id",
    "PrimaryContentManager/Title",

    "TeamName/TeamName"
  )
  .expand(
    "PrimaryContentManager",
    "Approvers",
    "NewsCreators",
    "BroadcastCreators",
    "EventCreators",
    "BrandCreators",
    "TeamName"
  )
  // Temporarily removed for testing
 .filter("TeamName/TeamName eq 'IS-Dummy'")
  .top(50)();

console.log("Current User:", currentUser);
console.log("SBU:", sbu);
console.log("Department:", profile.department);
console.log("Generated Filter Query:", filterQuery);
console.log("TeamCMList Items:", items);
console.log("Items Count:", items.length);
 

    if (!items.length) {console.log("FAILED: No TeamCMList records matched the filter");
    return false;
    }

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );
console.log("Final Items:", finalItems);
console.log("Final Items Count:", finalItems.length);

if (!finalItems.length) {
  console.log("FAILED: applyTeamPriority returned no records");
  return false;
}

    for (const item of finalItems) {
      //  Check if type enabled
      const isTypeEnabled =
        (type === "News" && item.IsNewsCreator === true) ||
        (type === "BroadCast" && item.IsBroadcastCreator === true) ||
        (type === "Event" && item.IsEventCreator === true) ||
        (type === "Brand" && item.IsBrandCreator === true);

      if (!isTypeEnabled) continue;


      const primaryCM = this.normalizeUsers(item.PrimaryContentManager);
      const approvers = this.normalizeUsers(item.Approvers);

      const creators =
        type === "News"
          ? this.normalizeUsers(item.NewsCreators)
          : type === "BroadCast"
            ? this.normalizeUsers(item.BroadcastCreators)
            : type === "Event"
              ? this.normalizeUsers(item.EventCreators)
              : this.normalizeUsers(item.BrandCreators);

              
      console.log("Team Item:", item.Id);
console.log("Requested Type:", type);
console.log("Is Type Enabled:", isTypeEnabled);
console.log("Primary CM:", primaryCM);
console.log("Approvers:", approvers);
console.log("Creators:", creators);
console.log("Current User:", currentUser);

      const isPrimaryOrApprover =
        primaryCM.some((u) => this.isSameUser(u, currentUser)) ||
        approvers.some((u) => this.isSameUser(u, currentUser));

      if (isPrimaryOrApprover) return true;

      // const isExplicitCreator = creators.some((u) =>
      //   this.isSameUser(u, currentUser),
      // );
      
     // let isExplicitCreator = false;

      // for (const u of creators) {
      //   // Direct user
      //   if (
      //     u.Id === currentUser.Id ||
      //     u.EMail?.toLowerCase() === currentUser.Email?.toLowerCase()
      //   ) {
      //     isExplicitCreator = true;
      //     break;
      //   }

      //   // Group match
      //   if (u.Title && userGroups.has(u.Title.trim().toLowerCase())) {
      //     isExplicitCreator = true;
      //     break;
      //   }
      // }
      
const isExplicitCreator = creators.some((u: any) => {
  const isDirectUser = this.isSameUser(u, currentUser);

  const isGroupMember =
    !!u.Title &&
    userGroups.has(u.Title.trim().toLowerCase());

  return isDirectUser || isGroupMember;
});

if (isExplicitCreator) return true;


      
    }

    return false;
  }

  // GET NEWS LIST DATA
  // public async getCurrentUserId(): Promise<number> {
  //   const me = await this.sp.web.currentUser();
  //   return me.Id;
  // }

  public async getCurrentUserId(): Promise<number> {
  const { currentUser } = await this.getUserContext();
  return currentUser.Id;
}

  public async getNewsForCreator(
    status: string,
    pageSize: number,
  ): Promise<any[]> {
    // const { currentUser, profile, sbu } = await this.getUserContext();
    // const userId = currentUser.Id;

const { currentUser, profile, sbu } = await this.getUserContext();
//const { currentUser} = await this.getUserContext();

const userId = Number(currentUser.Id);

//const userEmail = (currentUser.Email || "").trim().toLowerCase();
// console.log("Fetching news for user:", {
//   id: userId,
//   email: userEmail,
// });


    let filterQuery = `Created ge datetime'2024-12-31T00:00:00Z' and CommunicationType eq 'News'`;

    let teamName = (sbu || "").trim();
if (
      teamName === "HRS" ||
      teamName === "HRD" ||
      teamName === "HUMAN RESOURCES"
    ) {
      teamName = "HUMAN RESOURCES";
    }

    if (teamName === "HUMAN RESOURCES") {
      filterQuery += ` and (SBU eq 'HRS' or SBU eq 'HRD' or SBU eq 'HUMAN RESOURCES')`;
    } else if (teamName) {
      filterQuery += ` and SBU eq '${teamName.replace(/'/g, "''")}'`;
    } else if (profile.department) {
      filterQuery += ` and Department eq '${profile.department.replace(/'/g, "''")}'`;
    }
   

    if (status !== "All") {
      if (status === "Draft") {
        filterQuery += `
      and Status eq 'Draft'
      and AuthorId eq ${userId}
    `;
      } else if (status === "Pending") {
        filterQuery += `
      and PendingWithId eq ${userId}
      and (
        Status eq 'PendingWithLM'
        or Status eq 'PendingWithCM'
        or Status eq 'PendingWithCreator'
      )
    `;
      } else if (status === "Approval Queue") {
        filterQuery += `
      and AuthorId eq ${userId}
      and (
        Status eq 'PendingWithLM'
        or Status eq 'PendingWithCM'
      ) 
    `;
      } else if (status === "Published") {
        filterQuery += `
      and Status eq 'Published'
      and (
        AuthorId eq ${userId}
        or AssignToId eq ${userId}
        or FollowOnActionById eq ${userId}
        or PendingWithId eq ${userId}
      )
    `;
      } else if (status === "View Requests") {
        filterQuery += `
    and (
      (
        AuthorId eq ${userId} and (
          Status eq 'Draft'
          or Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'PendingWithCreator'
          or Status eq 'Published'
          or Status eq 'Rejected'
        )
      )
      or
      (
        AssignToId eq ${userId} and (
          Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithCreator'

        )
      )
      or
      (
        FollowOnActionById eq ${userId} and (
          Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithLM'

        )
      )
      or
      (
        PendingWithId eq ${userId} and (
          Status eq 'PendingWithLM'
          or Status eq 'PendingWithCM'
          or Status eq 'Published'
          or Status eq 'Rejected'
          or Status eq 'PendingWithCreator'

        )
      )
    )
  `;
      }
    }

    const items = await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.select(
        "Id",
        "Title",
        "SBU",
        "CommunicationType",
        "MainDescription",
        "PublishIn",

        "ExpiryDate",
        "EventDate",
        "NewsTypes",

        "NewsTypes/Label",
        "NewsTypes/TermGuid",

        "AssignToId",
        "PendingWithId",
        "FollowOnActionById",
        "ViewById",

        "InformToBrand",
        "IsReassigned",

        "Thumbnail",
        "ThumbnailCaption",

        "Picture1",
        "Pic1Caption",
        "Picture2",
        "Pic2Caption",
        "Picture3",
        "Pic3Caption",

        "AdditionalInformation1",
        "AdditionalInformation2",
        "AdditionalInformation3",

        "Comments",

        "Created",
        "Status",
        "PublishedDate",

        "AuthorId",
        "Author/Title",
        "PendingWith/Id",
        "PendingWith/EMail",
        "PendingWith/Title",
        "AttachmentFiles/FileName",
        "AttachmentFiles/ServerRelativeUrl",
      )
      .expand("Author", "PendingWith", "AttachmentFiles")
      .filter(filterQuery)
      .orderBy("Created", false)
      .top(pageSize)();

      
console.log("========== PENDING USER ID DIAGNOSTICS ==========");
console.log("Logged-in SharePoint user:", {
  id: Number(currentUser.Id),
  email: currentUser.Email,
  title: currentUser.Title,
});

console.log("Requested tab:", status);
console.log("Applied OData filter:", filterQuery);
console.log("Items returned:", items.length);

items.forEach((item: any) => {
  const currentUserId = Number(currentUser.Id);
  const pendingWithId = Number(item.PendingWithId);
  const isUserMatch =
    item.PendingWithId != null &&
    currentUserId === pendingWithId;

  console.log("[PendingWith ID Comparison]", {
    itemId: item.Id,
    title: item.Title,
    status: item.Status,

    // ID stored in the CorpCommunication Person field
    pendingWithId: item.PendingWithId,

    // Expanded SharePoint user details
    pendingUserId: item.PendingWith?.Id,
    pendingUserTitle: item.PendingWith?.Title,
    pendingUserEmail: item.PendingWith?.EMail,

    // Logged-in user's identity
    currentUserId,
    currentUserEmail: currentUser.Email,

    // Final ID comparison
    isUserMatch,
  });
});

console.log("========== END PENDING DIAGNOSTICS ==========");


    return items.map((it: any) => {
      const files = it.AttachmentFiles || [];

      const thumbFile = files.find(
        (f: any) =>
          f.FileName?.toLowerCase() === (it.Thumbnail || "").toLowerCase(),
      );

      let thumbUrl = "";

      if (thumbFile?.ServerRelativeUrl) {
        // force refresh to avoid browser cache
        thumbUrl = thumbFile.ServerRelativeUrl + `?v=${new Date().getTime()}`;
      }

      return {
        ...it,
        ThumbnailUrl: thumbUrl,
      };
    });
  }

  // Delete LIST ITEM

  // public async deleteNewsPermanently(id: number): Promise<void> {
  //   await this.sp.web.lists
  //     .getByTitle("CorpCommunication")
  //     .items.getById(id)
  //     .delete();
  // }

  // deleteing permanently
  public async deleteNewsPermanently(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .delete();

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      //  SUCCESS
      await this.auditService.log({
        action: "Delete",
        operation: "DELETE",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });
    } catch (error) {
      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");
      //  FAILURE
      await this.auditService.log({
        action: "Delete Failed",
        operation: "DELETE",
        statusCode: 500,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  // UPDATE LIST ITEM
  public async updateNews(id: number, payload: any) {
    try {
      await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(id)
        .update(payload);

      //  Determine action based on Status
      let actionName = "Update";

      const status = payload?.Status;
      const reassigned = payload?.IsReassigned ?? false;

      if (status === "PendingWithLM" && reassigned === false) {
        actionName = "Submit";
      } else if (status === "Rejected") {
        actionName = "Reject";
      } else if (status === "PendingWithCreator") {
        actionName = "Reassign";
      } else if (status === "PendingWithCM") {
        actionName = "Approve";
      } else if (status === "Published") {
        actionName = "Approve";
      } else if (status === "PendingWithLM" && reassigned === true) {
        actionName = "Reassign";
      }

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");
      //  SUCCESS LOG
      await this.auditService.log({
        action: actionName,
        operation: "POST",
        statusCode: 200,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });
    } catch (error) {
      ErrorService.logError("NewsService.updateNews", error);

      //  FAILURE LOG (also dynamic)
      let failAction = "Update Failed";

      const status = payload?.Status;
      const reassigned = payload?.IsReassigned ?? false;

      if (status === "PendingWithLM" && reassigned === false) {
        failAction = "Submit Failed";
      } else if (status === "Rejected") {
        failAction = "Reject Failed";
      } else if (status === "PendingWithCreator") {
        failAction = "Reassign Failed";
      } else if (status === "PendingWithCM") {
        failAction = "Approve Failed";
      } else if (status === "PendingWithLM" && reassigned === true) {
        failAction = "Reassign Failed";
      } else if (status === "Published") {
        failAction = "Approve Failed";
      }

      const roles = await this.getCurrentUserRoles();
      const roleString = roles.join(", ");

      await this.auditService.log({
        action: failAction,
        operation: "POST",
        statusCode: 500,
        listName: "CorpCommunication",
        itemId: id,
        userRole: roleString,
      });

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  // upload and delete the attchmentts
  public async uploadAttachments(itemId: number, files: File[]) {
    try {
      const item = this.sp.web.lists
        .getByTitle("CorpCommunication")
        .items.getById(itemId);

      const existingFiles = await item.attachmentFiles();

      for (const file of files) {
        const alreadyExists = existingFiles.find(
          (f) => f.FileName === file.name,
        );

        if (alreadyExists) {
          await item.attachmentFiles.getByName(file.name).delete();

          await new Promise((r) => setTimeout(r, 200));
        }

        await item.attachmentFiles.add(file.name, file);
      }
    } catch (error) {
      ErrorService.logError("NewsService.uploadAttachments", error);

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  // Get attachments
  public async getAttachmentLinks(
    itemId: number,
  ): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
    const files = await this.sp.web.lists
      .getByTitle("CorpCommunication")
      .items.getById(itemId)
      .attachmentFiles();

    return files.map((f: any) => ({
      FileName: f.FileName,
      ServerRelativeUrl: f.ServerRelativeUrl,
    }));
  }
  //  People picker helper
  public async getUserEmailById(userId: number): Promise<string> {
    const user = await this.sp.web.siteUsers.getById(userId)();
    return user?.Email || "";
  }

  // TermSetId
  public async getTermSetId(fieldInternalName: string): Promise<string> {
    try {
      const field: any = await this.sp.web.lists
        .getByTitle("CorpCommunication")
        .fields.getByInternalNameOrTitle(fieldInternalName)();

      return field.TermSetId;
    } catch (error) {
      ErrorService.logError("NewsService.getTermSetId", error);

      throw new Error(ErrorService.getUserMessage(error));
    }
  }

  //  taxonomy helper
  // public getTaxonomyValue(term: any) {
  //   if (!term) return null;

  //   return {
  //     Label: term.Label,
  //     TermGuid: term.TermGuid,
  //     WssId: -1,
  //   };
  // }
  public getTaxonomyValue(term: any) {
    if (!term) return null;

    return {
      Label: term.Label || term.text,
      TermGuid: term.TermGuid || term.key,
      WssId: -1,
    };
  }

  public async getTermSetTerms(
    termSetId: string,
  ): Promise<{ key: string; text: string }[]> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/v2.1/termStore/sets/${termSetId}/terms?$select=id,labels`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    if (!response.ok) {
      ErrorService.logError("NewsService.getTermSetTerms", {
        status: response.status,
      });

      throw new Error(
        ErrorService.getUserMessage({
          status: response.status,
        }),
      );
    }
    const json = await response.json();

    // json.value contains terms
    return (json.value || []).map((t: any) => ({
      key: t.id,
      text: t.labels?.[0]?.name || t.id,
    }));
  }

  // role check
  public async isCurrentUserApprover(): Promise<boolean> {
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);
    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "Approvers/Id",
        // "Approvers/EMail",
        "TeamName/TeamName",
      )
      .expand("Approvers", "TeamName")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    return finalItems.some((row: any) =>
      row.Approvers?.some((u: any) => this.isSameUser(u, currentUser)),
    );
  }

  public async isCurrentUserPrimaryContentManager(): Promise<boolean> {
    const { currentUser, profile, sbu } = await this.getUserContext();
    const filterQuery = this.buildTeamFilter(profile, sbu);
    const items = await this.sp.web.lists
      .getByTitle("TeamCMList")
      .items.select(
        "Id",
        "PrimaryContentManager/Id",
        // "PrimaryContentManager/EMail",
        "TeamName/TeamName",
      )
      .expand("PrimaryContentManager", "TeamName")
      .filter(filterQuery)();

    const finalItems = this.applyTeamPriority(
      items,
      sbu?.trim(),
      profile.department?.trim(),
      currentUser,
    );

    return finalItems.some((row: any) =>
      row.PrimaryContentManager?.some((u: any) =>
        this.isSameUser(u, currentUser),
      ),
    );
  }

  public async createFolderIfNotExists(
    libraryName: string,
    folderName: string,
  ): Promise<void> {
    try {
      const serverUrl = this.context.pageContext.web.serverRelativeUrl;

      const folderPath = `${serverUrl}/${libraryName}/${folderName}`;

      await this.sp.web.getFolderByServerRelativePath(folderPath)();
    } catch (error) {
      try {
        await this.sp.web.lists
          .getByTitle(libraryName)
          .rootFolder.folders.addUsingPath(folderName);
      } catch (createError) {
        ErrorService.logError(
          "NewsService.createFolderIfNotExists",
          createError,
        );

        throw new Error(ErrorService.getUserMessage(createError));
      }
    }
  }

  // getting files from transferlist images library data
  public async getTransferImagesByCorpId(corpId: number): Promise<any[]> {
    const folderPath = `${this.context.pageContext.web.serverRelativeUrl}/TransferListImages/${corpId}`;

    return await this.sp.web
      .getFolderByServerRelativePath(folderPath)
      .files.select("Name", "ServerRelativeUrl")();
  }

  // getting files from transfersbuindexlist data
  public async getTransferSbuIndexesByCorpId(corpId: number): Promise<any[]> {
    const listTitle = "TransferSbuIndexList";
    const folderTitle = corpId.toString();

    // 1️ Get real folder path (FileRef) using Title
    const folderItems = await this.sp.web.lists
      .getByTitle(listTitle)
      .items.filter(`Title eq '${folderTitle}' and FSObjType eq 1`)
      .select("FileRef")();

    if (!folderItems || folderItems.length === 0) {
      return [];
    }

    const realFolderPath = folderItems[0].FileRef;

    // 2️ Get items inside that folder using real internal path
    const items = await this.sp.web.lists
      .getByTitle(listTitle)
      .items.select("Id", "StartIndex", "SBU/Id", "SBU/TeamName", "FileDirRef")
      .expand("SBU")
      .filter(`startswith(FileDirRef,'${realFolderPath}')`)
      .orderBy("Id", true)();

    return items;
  }

  public async getTermNameFromHiddenList(guid: string): Promise<string> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("TaxonomyHiddenList")
        .items.filter(`IdForTerm eq '${guid}'`)
        .select("Term")
        .top(1)();

      return items?.[0]?.Term || "";
    } catch (error) {
      ErrorService.logError("NewsService.getTermNameFromHiddenList", error);

      return "";
    }
  }
}
