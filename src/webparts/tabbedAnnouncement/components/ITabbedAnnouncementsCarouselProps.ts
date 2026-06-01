import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITabbedAnnouncementsCarouselProps {
  webpartTitle: string;
  sourceList: string;
  highlightTypes: string;
  sites: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  currentUserLogin: string;
  currentUserId: number;
  isAdmin: boolean;
}
