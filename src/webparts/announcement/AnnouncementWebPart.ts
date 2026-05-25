import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AnnouncementWebPartStrings';
import AnnouncementsCarousel from './components/AnnouncementsCarousel';
import { IAnnouncementsCarouselProps } from './components/IAnnouncementsCarouselProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";

export interface IAnnouncementWebPartProps {
  description: string;
  sourceList: string;
  carouselLimit: number;
}

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementWebPartProps> {

  private _sp!: ReturnType<typeof spfi>;
  private _siteLists: { key: string; text: string }[] = [];
  private _currentUserId: number = 0;

  public render(): void {
    const element: React.ReactElement<IAnnouncementsCarouselProps> = React.createElement(
      AnnouncementsCarousel,
      {
        description: this.properties.description,
        sourceList: this.properties.sourceList,
        // Cap at 5 at the webpart level; component also enforces this
        carouselLimit: Math.min(this.properties.carouselLimit || 3, 5),
        isDarkTheme: !!this.context.pageContext.legacyPageContext?.isDarkTheme,
        environmentMessage: '',
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        currentUserLogin: this.context.pageContext.user.loginName,
        currentUserId: this._currentUserId,
        isAdmin: false,
      } as IAnnouncementsCarouselProps
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    await Promise.all([
      this._loadSiteLists(),
      this._loadCurrentUser(),
    ]);
    return super.onInit();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (newValue !== oldValue) {
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _loadCurrentUser(): Promise<void> {
    try {
      const currentUser = await this._sp.web.currentUser.select("Id")();
      this._currentUserId = currentUser.Id;
    } catch (err) {
      console.error("Failed to load current user:", err);
    }
  }

  private async _loadSiteLists(): Promise<void> {
    try {
      const lists = await this._sp.web.lists
        .filter("BaseTemplate eq 100 and Hidden eq false")
        .select("Title")();
      this._siteLists = lists.map(l => ({ key: l.Title, text: l.Title }));
    } catch (err) {
      console.error("Failed to load site lists:", err);
      this._siteLists = [];
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneDropdown('sourceList', {
                  label: 'Select Announcements List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneSlider('carouselLimit', {
                  label: 'Featured Announcements Limit',
                  min: 1,
                  max: 5,
                  step: 1,
                  showValue: true,
                  value: Math.min(this.properties.carouselLimit || 3, 5),
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
