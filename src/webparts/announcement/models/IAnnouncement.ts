export interface IAnnouncementAuthor {
  Id: number;
  Title: string;
}

export interface IAnnouncementAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IAnnouncementAudience {
  Id: number;
  Title: string;
}

export interface IAnnouncement {
  Id: number;
  Title: string;
  Body?: string;
  AnnouncementType: string;
  PublishDate: Date;
  ExpiryDate?: Date;
  Priority: "Critical" | "High" | "Medium" | "Low";
  Status: "Draft" | "Scheduled" | "Active" | "Expired";
  TargetAudienceType?: "All" | "Specific" | "Except";
  TargetAudience?: IAnnouncementAudience[];
  BannerImageUrl?: string;
  Attachments?: IAnnouncementAttachment[];
  Author?: IAnnouncementAuthor;
  Created?: Date;
}