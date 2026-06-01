export interface ITabbedAnnouncementAuthor {
  Id: number;
  Title: string;
}

export interface ITabbedAnnouncementAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface ITabbedAnnouncementAudience {
  Id: number;
  Title: string;
}

export interface ITabbedAnnouncement {
  Id: number;
  Title: string;
  Body?: string;
  HighlightType: string;
  Site: string;
  Priority: "Critical" | "High" | "Medium" | "Low";
  Status: "Draft" | "Published";
  TargetAudienceType?: "All" | "Specific" | "Except";
  TargetAudience?: ITabbedAnnouncementAudience[];
  BannerImageUrl?: string;
  Attachments?: ITabbedAnnouncementAttachment[];
  Author?: ITabbedAnnouncementAuthor;
  Created?: Date;
}
