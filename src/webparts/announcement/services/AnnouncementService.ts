import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import { IAnnouncement, IAnnouncementAttachment, IAnnouncementAudience } from "../models/IAnnouncement";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IAnnouncementListItem {
  Id: number;
  Title: string;
  Body?: string;
  AnnouncementType: string;
  PublishDate: string;
  ExpiryDate?: string;
  Priority: string;
  Status: string;
  TargetAudienceType?: string;
  TargetAudience?: IAnnouncementAudience[];
  BannerImageUrl?: { Url: string; Description?: string };
  Author?: { Id: number; Title: string };
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
  Created: string;
}

interface IAnnouncementItemData {
  Title: string;
  Body: string;
  AnnouncementType: string;
  PublishDate: string | undefined;
  ExpiryDate?: string | undefined;
  Priority: string;
  Status: string;
  TargetAudienceType?: string;
  TargetAudienceId?: number[];
  BannerImageUrl?: { Url: string; Description: string } | undefined;
}

const PRIORITY_ORDER: Record<string, number> = {
  Critical: 0,
  High: 1,
  Medium: 2,
  Low: 3,
};

export class AnnouncementService {
  private listName: string;
  private sp: SPFI;

  constructor(context: WebPartContext, listName: string) {
    this.listName = listName;
    this.sp = spfi().using(SPFx(context));
  }

  public async getAnnouncements(currentUserLogin: string): Promise<IAnnouncement[]> {
    const now = new Date();

    let currentUserTitle = "";
    try {
      const currentUser = await this.sp.web.currentUser.select("Title")();
      currentUserTitle = currentUser.Title?.toLowerCase().trim() ?? "";
    } catch {
      console.warn("Could not fetch current user title");
    }

    const items: IAnnouncementListItem[] = await this.sp.web.lists
      .getByTitle(this.listName)
      .items.select(
        "Id", "Title", "Body", "AnnouncementType",
        "PublishDate", "ExpiryDate", "Priority", "Status",
        "TargetAudienceType", "TargetAudience/Id", "TargetAudience/Title",
        "BannerImageUrl", "Author/Id", "Author/Title",
        "AttachmentFiles", "Created"
      )
      .expand("TargetAudience", "Author", "AttachmentFiles")();

    const statusUpdates: Promise<void>[] = [];

    for (const i of items) {
    // Critical: force Active immediately
    if (i.Priority?.trim() === "Critical" && i.Status?.trim() !== "Active" && i.Status?.trim() !== "Expired") {
        statusUpdates.push(
        this.sp.web.lists
            .getByTitle(this.listName)
            .items.getById(i.Id)
            .update({ Status: "Active" })
            .then(() => { i.Status = "Active"; })
        );
    }

    // Scheduled → Active when PublishDate is hit
    if (i.Status?.trim() === "Scheduled" && new Date(i.PublishDate) <= now) {
        statusUpdates.push(
        this.sp.web.lists
            .getByTitle(this.listName)
            .items.getById(i.Id)
            .update({ Status: "Active" })
            .then(() => { i.Status = "Active"; })
        );
    }

    // Active → Expired when ExpiryDate is passed
    if (i.Status?.trim() === "Active" && i.ExpiryDate && new Date(i.ExpiryDate) < now) {
        statusUpdates.push(
        this.sp.web.lists
            .getByTitle(this.listName)
            .items.getById(i.Id)
            .update({ Status: "Expired" })
            .then(() => { i.Status = "Expired"; })
        );
    }
    }

    await Promise.all(statusUpdates);

    const announcements: IAnnouncement[] = items.map((i) => ({
      Id: i.Id,
      Title: i.Title,
      Body: i.Body,
      AnnouncementType: i.AnnouncementType,
      PublishDate: new Date(i.PublishDate),
      ExpiryDate: i.ExpiryDate ? new Date(i.ExpiryDate) : undefined,
      Priority: i.Priority as "Critical" | "High" | "Medium" | "Low",
      Status: i.Status as "Draft" | "Scheduled" | "Active" | "Expired",
      TargetAudienceType: (i.TargetAudienceType || "All") as "All" | "Specific" | "Except",
      TargetAudience: i.TargetAudience ?? [],
      BannerImageUrl: i.BannerImageUrl?.Url,
      Author: i.Author ? { Id: i.Author.Id, Title: i.Author.Title } : undefined,
      Attachments: (i.AttachmentFiles ?? []).map((f) => ({
        FileName: f.FileName,
        ServerRelativeUrl: f.ServerRelativeUrl,
      } as IAnnouncementAttachment)),
      Created: new Date(i.Created),
    }));

    // Target Audience filtering
    const filtered = announcements.filter((a) => {
      const type = a.TargetAudienceType || "All";
      if (type === "All") return true;
      const audienceTitles = (a.TargetAudience || []).map((p) => p.Title?.toLowerCase().trim());
      const isInAudience = audienceTitles.includes(currentUserTitle);
      if (type === "Specific") return isInAudience;
      if (type === "Except") return !isInAudience;
      return true;
    });

    // Sort: Priority → PublishDate → Created
    filtered.sort((a, b) => {
      const priorityDiff = PRIORITY_ORDER[a.Priority] - PRIORITY_ORDER[b.Priority];
      if (priorityDiff !== 0) return priorityDiff;
      const publishDiff = a.PublishDate.getTime() - b.PublishDate.getTime();
      if (publishDiff !== 0) return publishDiff;
      return (a.Created?.getTime() ?? 0) - (b.Created?.getTime() ?? 0);
    });

    return filtered;
  }

  public async addAnnouncement(
    announcement: Partial<IAnnouncement>,
    bannerFile?: File,
    attachments?: File[]
  ): Promise<void> {
    let bannerUrl: string | undefined;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    // Critical always goes Active immediately
    const resolvedStatus = announcement.Priority === "Critical" ? "Active" : (announcement.Status || "Draft");

    const itemData: IAnnouncementItemData = {
      Title: announcement.Title ?? "",
      Body: announcement.Body || "",
      AnnouncementType: announcement.AnnouncementType ?? "",
      PublishDate: announcement.PublishDate?.toISOString(),
      ExpiryDate: announcement.ExpiryDate?.toISOString() ?? undefined,
      Priority: announcement.Priority ?? "Medium",
      Status: resolvedStatus,
      TargetAudienceType: announcement.TargetAudienceType || "All",
    };

    if (announcement.TargetAudience && announcement.TargetAudience.length > 0) {
      const audienceIds = announcement.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: announcement.Title || "Announcement Banner" };
    }

    const addResult = await this.sp.web.lists.getByTitle(this.listName).items.add(itemData);

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(addResult.Id, attachments);
    }
  }

  public async updateAnnouncement(
    announcementId: number,
    announcement: Partial<IAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> {
    let bannerUrl = announcement.BannerImageUrl;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    // Critical always goes Active immediately
    const resolvedStatus = announcement.Priority === "Critical" ? "Active" : (announcement.Status || "Draft");

    const itemData: IAnnouncementItemData = {
      Title: announcement.Title ?? "",
      Body: announcement.Body || "",
      AnnouncementType: announcement.AnnouncementType ?? "",
      PublishDate: announcement.PublishDate?.toISOString(),
      ExpiryDate: announcement.ExpiryDate?.toISOString() ?? undefined,
      Priority: announcement.Priority ?? "Medium",
      Status: resolvedStatus,
      TargetAudienceType: announcement.TargetAudienceType || "All",
    };

    if (announcement.TargetAudience && announcement.TargetAudience.length > 0) {
      const audienceIds = announcement.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: announcement.Title || "Announcement Banner" };
    } else if (!bannerFile && !announcement.BannerImageUrl) {
      itemData.BannerImageUrl = undefined;
    }

    await this.sp.web.lists.getByTitle(this.listName).items.getById(announcementId).update(itemData);

    if (deletedAttachmentNames && deletedAttachmentNames.length > 0) {
      await this.deleteAttachments(announcementId, deletedAttachmentNames);
    }

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(announcementId, attachments);
    }
  }

  public async uploadAttachments(itemId: number, files: File[]): Promise<void> {
    for (const file of files) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.add(file.name, file);
      } catch (error) {
        console.error("Failed to upload attachment " + file.name + ":", error);
      }
    }
  }

  public async deleteAttachments(itemId: number, fileNames: string[]): Promise<void> {
    for (const fileName of fileNames) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.getByName(fileName)
          .delete();
      } catch (error) {
        console.error("Failed to delete attachment " + fileName + ":", error);
      }
    }
  }

  private async uploadBanner(file: File): Promise<string> {
    const folderPath = "SiteAssets/AnnouncementsImages";
    const fileName = Date.now().toString() + "_" + file.name;
    try {
      try {
        await this.sp.web.getFolderByServerRelativePath(folderPath).select("Exists")();
      } catch {
        await this.sp.web.folders.addUsingPath(folderPath);
      }
      const uploadResult = await this.sp.web
        .getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(fileName, file, { Overwrite: true });
      return uploadResult.ServerRelativeUrl;
    } catch (error) {
      console.error("Banner upload error:", error);
      throw new Error("Failed to upload banner image");
    }
  }
}