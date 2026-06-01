import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TabbedAnnouncementWebPartStrings';
import TabbedAnnouncementsCarousel from './components/TabbedAnnouncementsCarousel';
import { ITabbedAnnouncementsCarouselProps } from './components/ITabbedAnnouncementsCarouselProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

export interface ITabbedAnnouncementWebPartProps {
  webpartTitle: string;
  sourceList: string;
  highlightTypes: string;
  sites: string;
  editGroup: string;
}

export default class TabbedAnnouncementWebPart extends BaseClientSideWebPart<ITabbedAnnouncementWebPartProps> {

  private _sp!: ReturnType<typeof spfi>;
  private _siteLists: { key: string; text: string }[] = [];
  private _currentUserId: number = 0;
  private _isEditor: boolean = false;

  public render(): void {
    const element: React.ReactElement<ITabbedAnnouncementsCarouselProps> = React.createElement(
      TabbedAnnouncementsCarousel,
      {
        webpartTitle: this.properties.webpartTitle || 'MGEN Thermal Highlights',
        sourceList: this.properties.sourceList,
        highlightTypes: this.properties.highlightTypes || '',
        sites: this.properties.sites || '',
        isDarkTheme: !!this.context.pageContext.legacyPageContext?.isDarkTheme,
        environmentMessage: '',
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        currentUserLogin: this.context.pageContext.user.loginName,
        currentUserId: this._currentUserId,
        isAdmin: this._isEditor,
      } as ITabbedAnnouncementsCarouselProps
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
      this._checkEditPermission(),
    ]);
    return super.onInit();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (newValue !== oldValue) {
      if (propertyPath === 'editGroup') {
        // Re-evaluate group membership when the group name changes
        this._checkEditPermission().then(() => this.render()).catch(console.error);
      } else {
        this.render();
      }
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

  private async _checkEditPermission(): Promise<void> {
    const groupName = this.properties.editGroup?.trim();
    if (!groupName) {
      // No group configured — no one gets editor access by default
      this._isEditor = false;
      return;
    }
    try {
      const userGroups = await this._sp.web.currentUser.groups();
      this._isEditor = userGroups.some(g => g.Title === groupName);
    } catch (err) {
      console.error("Failed to check edit group membership:", err);
      this._isEditor = false;
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
                PropertyPaneTextField('webpartTitle', {
                  label: 'Web Part Title',
                  placeholder: 'MGEN Thermal Highlights',
                }),
                PropertyPaneDropdown('sourceList', {
                  label: 'Select Highlights List',
                  options: this._siteLists,
                  disabled: this._siteLists.length === 0,
                }),
                PropertyPaneTextField('highlightTypes', {
                  label: 'Highlight Types',
                  description: 'Comma-separated (e.g. Achievement, Certification, Recognition)',
                  placeholder: 'Achievement, Certification, Recognition',
                }),
                PropertyPaneTextField('sites', {
                  label: 'Sites',
                  description: 'Comma-separated — each becomes a tab (e.g. Corporate, Panay, Cebu)',
                  placeholder: 'Corporate, Panay, Cebu',
                }),
                PropertyPaneTextField('editGroup', {
                  label: 'Add/Edit Permission Group',
                  description: 'SharePoint site group name — only members can add or edit highlights',
                  placeholder: 'e.g. MThermalEA',
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
