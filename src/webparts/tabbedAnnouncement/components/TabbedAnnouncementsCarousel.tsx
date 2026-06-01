import * as React from "react";
import { useEffect, useState } from "react";
import { Icon } from "@fluentui/react";
import styles from "./TabbedAnnouncementsCarousel.module.scss";
import { ITabbedAnnouncementsCarouselProps } from "./ITabbedAnnouncementsCarouselProps";
import { TabbedAnnouncementService } from "../services/TabbedAnnouncementService";
import { ITabbedAnnouncement } from "../models/ITabbedAnnouncement";
import TabbedAnnouncementDetailsPanel from "./TabbedAnnouncementDetailsPanel";
import ViewAllPanel from "./ViewAllPanel";
import AddEditTabbedAnnouncement from "./AddEditTabbedAnnouncement";

const ITEMS_PER_TAB = 5;

const TYPE_COLORS = [
  '#2e7d32',
  '#1565c0',
  '#6a1b9a',
  '#e65100',
  '#00838f',
  '#ad1457',
  '#c62828',
  '#37474f',
];

const stripHtml = (html: string): string =>
  html.replace(/<[^>]+>/g, '').trim();

const TabbedAnnouncementsCarousel = (props: ITabbedAnnouncementsCarouselProps): JSX.Element => {
  const [highlights, setHighlights] = useState<ITabbedAnnouncement[]>([]);
  const [activeTab, setActiveTab] = useState<string>('All');
  const [selectedHighlight, setSelectedHighlight] = useState<ITabbedAnnouncement | undefined>(undefined);
  const [isDetailsOpen, setIsDetailsOpen] = useState(false);
  const [isViewAllOpen, setIsViewAllOpen] = useState(false);
  const [isAddEditOpen, setIsAddEditOpen] = useState(false);
  const [addEditMode, setAddEditMode] = useState<'add' | 'edit'>('add');
  const [highlightToEdit, setHighlightToEdit] = useState<ITabbedAnnouncement | undefined>(undefined);

  const highlightTypesList = props.highlightTypes
    ? props.highlightTypes.split(',').map(t => t.trim()).filter(Boolean)
    : [];

  const sitesList = props.sites
    ? props.sites.split(',').map(s => s.trim()).filter(Boolean)
    : [];

  const tabs = ['All', ...sitesList];

  const getTypeColor = (type: string): string => {
    const idx = highlightTypesList.findIndex(t => t.toLowerCase() === type?.trim().toLowerCase());
    return idx >= 0 ? TYPE_COLORS[idx % TYPE_COLORS.length] : TYPE_COLORS[0];
  };

  const loadHighlights = async (): Promise<void> => {
    try {
      const service = new TabbedAnnouncementService(props.context, props.sourceList);
      const data = await service.getTabbedAnnouncements();
      setHighlights(data);
    } catch (err) {
      console.error('Failed to load highlights', err);
    }
  };

  useEffect(() => {
    if (!props.sourceList) return;
    loadHighlights().catch(console.error);
  }, [props.sourceList, props.context]);

  useEffect(() => {
    if (!tabs.includes(activeTab)) setActiveTab('All');
  }, [props.sites]);

  const visibleHighlights = highlights
    .filter(h => h.Status?.trim() === 'Published')
    .filter(h => activeTab === 'All' || h.Site?.trim() === 'All' || h.Site?.trim() === activeTab)
    .slice(0, ITEMS_PER_TAB);

  const formatDate = (date: Date): string =>
    date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const openDetails = (highlight: ITabbedAnnouncement): void => {
    setSelectedHighlight(highlight);
    setIsDetailsOpen(true);
  };

  const closeDetails = (): void => {
    setIsDetailsOpen(false);
    setSelectedHighlight(undefined);
  };

  const openEdit = (highlight: ITabbedAnnouncement): void => {
    setIsDetailsOpen(false);
    setAddEditMode('edit');
    setHighlightToEdit(highlight);
    setIsAddEditOpen(true);
  };

  const openAdd = (): void => {
    setAddEditMode('add');
    setHighlightToEdit(undefined);
    setIsAddEditOpen(true);
  };

  const closeAddEdit = (): void => {
    setIsAddEditOpen(false);
    if (highlightToEdit) {
      setSelectedHighlight(highlightToEdit);
      setIsDetailsOpen(true);
    }
    setHighlightToEdit(undefined);
  };

  const handleSave = async (
    data: Partial<ITabbedAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> => {
    const service = new TabbedAnnouncementService(props.context, props.sourceList);
    if (addEditMode === 'add') {
      await service.addTabbedAnnouncement(data, bannerFile, attachments);
    } else if (addEditMode === 'edit' && highlightToEdit) {
      await service.updateTabbedAnnouncement(highlightToEdit.Id, data, bannerFile, attachments, deletedAttachmentNames);
    }
    setIsAddEditOpen(false);
    setHighlightToEdit(undefined);
    setSelectedHighlight(undefined);
    setIsDetailsOpen(false);
    setIsViewAllOpen(false);
    await loadHighlights();
  };

  const renderPanels = (): JSX.Element => (
    <>
      <ViewAllPanel
        announcements={highlights}
        isOpen={isViewAllOpen}
        onDismiss={() => setIsViewAllOpen(false)}
        onSelectTabbedAnnouncement={openDetails}
        onAddTabbedAnnouncement={props.isAdmin ? openAdd : undefined}
        highlightTypesList={highlightTypesList}
      />
      <TabbedAnnouncementDetailsPanel
        announcement={selectedHighlight}
        isOpen={isDetailsOpen}
        onDismiss={closeDetails}
        onEdit={openEdit}
        currentUserId={props.currentUserId}
        isAdmin={props.isAdmin}
      />
      <AddEditTabbedAnnouncement
        isOpen={isAddEditOpen}
        mode={addEditMode}
        announcement={highlightToEdit}
        context={props.context}
        highlightTypeOptions={highlightTypesList}
        siteOptions={sitesList}
        onDismiss={closeAddEdit}
        onSave={handleSave}
      />
    </>
  );

  if (!props.sourceList) {
    return (
      <div className={styles.container}>
        <h2 className={styles.webpartTitle}>{props.webpartTitle || 'MGEN Thermal Highlights'}</h2>
        <p className={styles.noList}>Please select a list in the web part settings to get started.</p>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <h2 className={styles.webpartTitle}>{props.webpartTitle || 'MGEN Thermal Highlights'}</h2>

      {/* ── Tabs row ── */}
      <div className={styles.tabsRow}>
        <div className={styles.tabsLeft}>
          {tabs.map(tab => (
            <button
              key={tab}
              className={`${styles.tab}${activeTab === tab ? ' ' + styles.tabActive : ''}`}
              onClick={() => setActiveTab(tab)}
            >
              {tab}
            </button>
          ))}
        </div>
        <button className={styles.viewAllBtn} onClick={() => setIsViewAllOpen(true)}>
          View all
        </button>
      </div>

      {/* ── Items list ── */}
      {visibleHighlights.length === 0 ? (
        <div className={styles.emptyState}>
          <p>No active highlights{activeTab !== 'All' ? ` for ${activeTab}` : ''}.</p>
        </div>
      ) : (
        <ul className={styles.itemsList}>
          {visibleHighlights.map(highlight => {
            const bodyText = stripHtml(highlight.Body ?? '');
            return (
              <li key={highlight.Id} className={styles.item} onClick={() => openDetails(highlight)}>
                <div className={styles.itemIconBox}>
                  <Icon iconName="Document" />
                </div>
                <div className={styles.itemBody}>
                  <div
                    className={styles.itemCategory}
                    style={{ color: getTypeColor(highlight.HighlightType) }}
                  >
                    {highlight.HighlightType}
                  </div>
                  <div className={styles.itemTitle}>{highlight.Title}</div>
                  {bodyText && <div className={styles.itemDesc}>{bodyText}</div>}
                </div>
                <div className={styles.itemDate}>{highlight.Created ? formatDate(highlight.Created) : ''}</div>
              </li>
            );
          })}
        </ul>
      )}

      {renderPanels()}
    </div>
  );
};

export default TabbedAnnouncementsCarousel;
