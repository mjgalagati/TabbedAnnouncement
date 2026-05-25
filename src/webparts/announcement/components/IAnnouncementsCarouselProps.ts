import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAnnouncementsCarouselProps {
  description: string;
  sourceList: string;
  carouselLimit: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  currentUserLogin: string;
  currentUserId: number;
  isAdmin: boolean;
}