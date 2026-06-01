import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import { ITabbedAnnouncement, ITabbedAnnouncementAttachment, ITabbedAnnouncementAudience } from "../models/ITabbedAnnouncement";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface ITabbedAnnouncementListItem {
  Id: number;
  Title: string;
  Body?: string;
  HighlightType: string;
  Site: string;
  Priority: string;
  Status: string;
  TargetAudienceType?: string;
  TargetAudience?: ITabbedAnnouncementAudience[];
  BannerImageUrl?: { Url: string; Description?: string };
  Author?: { Id: number; Title: string };
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
  Created: string;
}

interface ITabbedAnnouncementItemData {
  Title: string;
  Body: string;
  HighlightType: string;
  Site: string;
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

export class TabbedAnnouncementService {
  private listName: string;
  private sp: SPFI;

  constructor(context: WebPartContext, listName: string) {
    this.listName = listName;
    this.sp = spfi().using(SPFx(context));
  }

  public async getTabbedAnnouncements(): Promise<ITabbedAnnouncement[]> {
    let currentUserTitle = "";
    try {
      const currentUser = await this.sp.web.currentUser.select("Title")();
      currentUserTitle = currentUser.Title?.toLowerCase().trim() ?? "";
    } catch {
      console.warn("Could not fetch current user title");
    }

    const items: ITabbedAnnouncementListItem[] = await this.sp.web.lists
      .getByTitle(this.listName)
      .items.select(
        "Id", "Title", "Body", "HighlightType", "Site",
        "Priority", "Status",
        "TargetAudienceType", "TargetAudience/Id", "TargetAudience/Title",
        "BannerImageUrl", "Author/Id", "Author/Title",
        "AttachmentFiles", "Created"
      )
      .expand("TargetAudience", "Author", "AttachmentFiles")();

    const highlights: ITabbedAnnouncement[] = items.map((i) => ({
      Id: i.Id,
      Title: i.Title,
      Body: i.Body,
      HighlightType: i.HighlightType,
      Site: i.Site,
      Priority: i.Priority as "Critical" | "High" | "Medium" | "Low",
      Status: i.Status as "Draft" | "Published",
      TargetAudienceType: (i.TargetAudienceType || "All") as "All" | "Specific" | "Except",
      TargetAudience: i.TargetAudience ?? [],
      BannerImageUrl: i.BannerImageUrl?.Url,
      Author: i.Author ? { Id: i.Author.Id, Title: i.Author.Title } : undefined,
      Attachments: (i.AttachmentFiles ?? []).map((f) => ({
        FileName: f.FileName,
        ServerRelativeUrl: f.ServerRelativeUrl,
      } as ITabbedAnnouncementAttachment)),
      Created: new Date(i.Created),
    }));

    const filtered = highlights.filter((h) => {
      const type = h.TargetAudienceType || "All";
      if (type === "All") return true;
      const audienceTitles = (h.TargetAudience || []).map((p) => p.Title?.toLowerCase().trim());
      const isInAudience = audienceTitles.includes(currentUserTitle);
      if (type === "Specific") return isInAudience;
      if (type === "Except") return !isInAudience;
      return true;
    });

    filtered.sort((a, b) => {
      const priorityDiff = PRIORITY_ORDER[a.Priority] - PRIORITY_ORDER[b.Priority];
      if (priorityDiff !== 0) return priorityDiff;
      return (b.Created?.getTime() ?? 0) - (a.Created?.getTime() ?? 0);
    });

    return filtered;
  }

  public async addTabbedAnnouncement(
    highlight: Partial<ITabbedAnnouncement>,
    bannerFile?: File,
    attachments?: File[]
  ): Promise<void> {
    let bannerUrl: string | undefined;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    const itemData: ITabbedAnnouncementItemData = {
      Title: highlight.Title ?? "",
      Body: highlight.Body || "",
      HighlightType: highlight.HighlightType ?? "",
      Site: highlight.Site ?? "",
      Priority: highlight.Priority ?? "Medium",
      Status: highlight.Status ?? "Draft",
      TargetAudienceType: highlight.TargetAudienceType || "All",
    };

    if (highlight.TargetAudience && highlight.TargetAudience.length > 0) {
      const audienceIds = highlight.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: highlight.Title || "Highlight Banner" };
    }

    const addResult = await this.sp.web.lists.getByTitle(this.listName).items.add(itemData);

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(addResult.Id, attachments);
    }
  }

  public async updateTabbedAnnouncement(
    highlightId: number,
    highlight: Partial<ITabbedAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> {
    let bannerUrl = highlight.BannerImageUrl;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    const itemData: ITabbedAnnouncementItemData = {
      Title: highlight.Title ?? "",
      Body: highlight.Body || "",
      HighlightType: highlight.HighlightType ?? "",
      Site: highlight.Site ?? "",
      Priority: highlight.Priority ?? "Medium",
      Status: highlight.Status ?? "Draft",
      TargetAudienceType: highlight.TargetAudienceType || "All",
    };

    if (highlight.TargetAudience && highlight.TargetAudience.length > 0) {
      const audienceIds = highlight.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: highlight.Title || "Highlight Banner" };
    } else if (!bannerFile && !highlight.BannerImageUrl) {
      itemData.BannerImageUrl = undefined;
    }

    await this.sp.web.lists.getByTitle(this.listName).items.getById(highlightId).update(itemData);

    if (deletedAttachmentNames && deletedAttachmentNames.length > 0) {
      await this.deleteAttachments(highlightId, deletedAttachmentNames);
    }

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(highlightId, attachments);
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
    const folderPath = "SiteAssets/TabbedAnnouncementImages";
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
