import * as React from 'react';
import { PrimaryButton, Icon, SearchBox, DatePicker, IconButton } from '@fluentui/react';
import { IAnnouncement } from '../models/IAnnouncement';
import styles from './ViewAllPanel.module.scss';
import { useState } from 'react';

interface ViewAllPanelProps {
  announcements: IAnnouncement[];
  isOpen: boolean;
  onDismiss: () => void;
  onSelectAnnouncement: (announcement: IAnnouncement) => void;
  onAddAnnouncement?: () => void;
}

const ViewAllPanel: React.FC<ViewAllPanelProps> = ({
  announcements, isOpen, onDismiss, onSelectAnnouncement, onAddAnnouncement,
}): JSX.Element => {
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [dateFrom, setDateFrom] = useState<Date | undefined>(undefined);
  const [dateTo, setDateTo] = useState<Date | undefined>(undefined);
  const [showDateFilter, setShowDateFilter] = useState<boolean>(false);

  if (!isOpen) return <></>;

  const filteredAnnouncements = announcements.filter(announcement => {
    const matchesSearch =
      announcement.Title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      announcement.AnnouncementType.toLowerCase().includes(searchQuery.toLowerCase()) ||
      announcement.Body?.toLowerCase().includes(searchQuery.toLowerCase());

    let matchesDateRange = true;
    if (dateFrom || dateTo) {
      const publishDate = new Date(announcement.PublishDate);
      if (dateFrom && dateTo) matchesDateRange = publishDate >= dateFrom && publishDate <= dateTo;
      else if (dateFrom) matchesDateRange = publishDate >= dateFrom;
      else if (dateTo) matchesDateRange = publishDate <= dateTo;
    }

    return matchesSearch && matchesDateRange;
  });

  const criticalAnnouncements  = filteredAnnouncements.filter(a => a.Priority?.trim() === 'Critical' && a.Status?.trim() === 'Active');
  const highAnnouncements      = filteredAnnouncements.filter(a => a.Priority?.trim() === 'High' && a.Status?.trim() === 'Active');
  const activeAnnouncements    = filteredAnnouncements.filter(a => (a.Priority?.trim() === 'Medium' || a.Priority?.trim() === 'Low') && a.Status?.trim() === 'Active');
  const scheduledAnnouncements = filteredAnnouncements.filter(a => a.Status?.trim() === 'Scheduled');
  const draftAnnouncements     = filteredAnnouncements.filter(a => a.Status?.trim() === 'Draft');
  const expiredAnnouncements   = filteredAnnouncements.filter(a => a.Status?.trim() === 'Expired');

  const formatDate = (date?: Date): string => {
    if (!date) return 'N/A';
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  const getStatusClass = (status: string): string => {
    if (status === 'Active')    return styles.statusActive;
    if (status === 'Scheduled') return styles.statusScheduled;
    if (status === 'Expired')   return styles.statusExpired;
    return styles.statusDraft;
  };

  const getPriorityClass = (priority: string): string => {
    if (priority === 'Critical') return styles.priorityCritical;
    if (priority === 'High')     return styles.priorityHigh;
    if (priority === 'Medium')   return styles.priorityMedium;
    return styles.priorityLow;
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const renderAnnouncementCard = (announcement: IAnnouncement): JSX.Element => (
    <div
      key={announcement.Id}
      className={styles.announcementCard}
      onClick={() => onSelectAnnouncement(announcement)}
    >
      {announcement.BannerImageUrl && (
        <div className={styles.cardBanner}>
          <img src={announcement.BannerImageUrl} alt={announcement.Title} />
        </div>
      )}
      <div className={styles.cardContent}>
        <div className={styles.cardHeader}>
          <h3 className={styles.cardTitle}>{announcement.Title}</h3>
          <div className={styles.cardBadges}>
            <span className={`${styles.priorityBadge} ${getPriorityClass(announcement.Priority)}`}>
              {announcement.Priority}
            </span>
            <span className={`${styles.statusBadge} ${getStatusClass(announcement.Status)}`}>
              {announcement.Status}
            </span>
          </div>
        </div>
        <div className={styles.cardInfo}>
          <div className={styles.infoItem}>
            <Icon iconName="Calendar" className={styles.infoIcon} />
            <span>{formatDate(announcement.PublishDate)}</span>
          </div>
          {announcement.ExpiryDate && (
            <div className={styles.infoItem}>
              <Icon iconName="EventDateMissed12" className={styles.infoIcon} />
              <span>Expires: {formatDate(announcement.ExpiryDate)}</span>
            </div>
          )}
          <div className={styles.infoItem}>
            <Icon iconName="Tag" className={styles.infoIcon} />
            <span className={styles.typeTag}>{announcement.AnnouncementType}</span>
          </div>
          {announcement.TargetAudienceType && announcement.TargetAudienceType !== 'All' && (
            <div className={styles.infoItem}>
              <Icon iconName="People" className={styles.infoIcon} />
              <span>{announcement.TargetAudienceType === 'Specific' ? 'Specific Audience' : 'All Except'}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <div className={styles.headerTop}>
            <h2 className={styles.panelTitle}>All Announcements</h2>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {onAddAnnouncement && (
                <PrimaryButton
                  text="Add Announcement"
                  iconProps={{ iconName: 'Add' }}
                  onClick={onAddAnnouncement}
                  className={styles.addButton}
                />
              )}
              <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
                <Icon iconName="Cancel" />
              </button>
            </div>
          </div>

          <div className={styles.filterBar}>
            <SearchBox
              placeholder="Search announcements..."
              onChange={(_, newValue) => setSearchQuery(newValue || '')}
              className={styles.searchBox}
            />
            <IconButton
              iconProps={{ iconName: showDateFilter ? 'ChevronUp' : 'Calendar' }}
              onClick={() => setShowDateFilter(!showDateFilter)}
              className={styles.dateFilterToggle}
              toggle
              checked={showDateFilter}
            />
          </div>

          {showDateFilter && (
            <div className={styles.dateFilters}>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>From:</label>
                <DatePicker
                  placeholder="Select start date"
                  value={dateFrom}
                  onSelectDate={(date) => setDateFrom(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  className={styles.datePicker}
                />
                {dateFrom && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateFrom(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>To:</label>
                <DatePicker
                  placeholder="Select end date"
                  value={dateTo}
                  onSelectDate={(date) => setDateTo(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  minDate={dateFrom}
                  className={styles.datePicker}
                />
                {dateTo && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateTo(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              {(dateFrom || dateTo) && (
                <IconButton
                  iconProps={{ iconName: 'ClearFilter' }}
                  onClick={() => { setDateFrom(undefined); setDateTo(undefined); }}
                  className={styles.clearAllButton}
                />
              )}
            </div>
          )}

          <div className={styles.announcementCount}>
            <span>{filteredAnnouncements.length} announcement{filteredAnnouncements.length !== 1 ? 's' : ''} found</span>
          </div>
        </div>

        {/* ── Scrollable Content ── */}
        <div className={styles.panelContent}>
          {filteredAnnouncements.length === 0 && (
            <div className={styles.emptyState}>
              <Icon iconName="Megaphone" className={styles.emptyIcon} />
              <p>No announcements found</p>
              <span>Try adjusting your search or add a new announcement</span>
            </div>
          )}

          {criticalAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Warning" className={styles.sectionIconCritical} />
                <h3>Critical ({criticalAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {criticalAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {highAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="SortUp" className={styles.sectionIconHigh} />
                <h3>High Priority ({highAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {highAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {activeAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="MegaphoneSolid" className={styles.sectionIcon} />
                <h3>Active ({activeAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {activeAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {scheduledAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="ScheduleEventAction" className={styles.sectionIcon} />
                <h3>Scheduled ({scheduledAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {scheduledAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {draftAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Edit" className={styles.sectionIcon} />
                <h3>Draft ({draftAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {draftAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {expiredAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="History" className={styles.sectionIcon} />
                <h3>Expired ({expiredAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {expiredAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}
        </div>

      </div>
    </div>
  );
};

export default ViewAllPanel;