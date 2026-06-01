import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, Icon } from '@fluentui/react';
import { ITabbedAnnouncement } from '../models/ITabbedAnnouncement';
import styles from './TabbedAnnouncementDetailsPanel.module.scss';

interface TabbedAnnouncementDetailsPanelProps {
  announcement?: ITabbedAnnouncement;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (announcement: ITabbedAnnouncement) => void;
  currentUserId: number;
  isAdmin: boolean;
}

const TabbedAnnouncementDetailsPanel: React.FC<TabbedAnnouncementDetailsPanelProps> = ({
  announcement, isOpen, onDismiss, onEdit, currentUserId, isAdmin,
}) => {
  if (!isOpen || !announcement) return null; // eslint-disable-line @rushstack/no-new-null

  const isAuthor = announcement.Author?.Id === currentUserId;
  const canEdit = isAdmin || isAuthor;

  const getFileIcon = (fileName: string): string => {
    const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
    if (ext === 'pdf') return 'PDF';
    if (ext === 'doc' || ext === 'docx') return 'WordDocument';
    if (ext === 'xls' || ext === 'xlsx' || ext === 'csv') return 'ExcelDocument';
    if (ext === 'ppt' || ext === 'pptx') return 'PowerPointDocument';
    if (ext === 'jpg' || ext === 'jpeg' || ext === 'png' || ext === 'gif' || ext === 'webp') return 'FileImage';
    if (ext === 'mp4') return 'Video';
    if (ext === 'mp3') return 'Music';
    return 'Attach';
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const hasBanner = !!announcement.BannerImageUrl;

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Hero Banner — title overlaid on image (article style) ── */}
        {hasBanner ? (
          <div className={styles.bannerContainer}>
            <img
              src={announcement.BannerImageUrl}
              alt={announcement.Title ?? 'Tabbed Announcement Banner'}
              className={styles.bannerImage}
            />

            {/* gradient + title/meta overlaid on banner */}
            <div className={styles.bannerOverlayContent}>
              <div className={styles.bannerTypeBadgeRow}>
                <span className={styles.bannerTypeBadge}>{announcement.HighlightType ?? 'General'}</span>
                <span className={styles.bannerTypeBadge}>{announcement.Priority}</span>
              </div>
              <h1 className={styles.bannerTitle}>{announcement.Title ?? 'Untitled'}</h1>
              {announcement.Author && (
                <span className={styles.bannerAuthor}>
                  <Icon iconName="Contact" /> {announcement.Author.Title}
                </span>
              )}
              {announcement.Created && (
                <span className={styles.bannerPublishDate}>
                  {announcement.Created.toLocaleDateString('en-US', {
                    weekday: 'short', year: 'numeric', month: 'short', day: 'numeric',
                  })}
                </span>
              )}
            </div>

            <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
              <Icon iconName="Cancel" />
            </button>
          </div>
        ) : (
          // No banner: close button anchored to top-right of modal
          <button
            className={`${styles.closeBtn} ${styles.closeBtnNoImage}`}
            onClick={onDismiss}
            aria-label="Close"
          >
            <Icon iconName="Cancel" />
          </button>
        )}

        {/* ── Scrollable article body ── */}
        <div className={styles.scrollBody}>
          <div className={styles.content}>

            {/* When no banner: show title/meta in content */}
            {!hasBanner && (
              <div className={styles.headerNoBanner}>
                <div className={styles.titleSection}>
                  <div className={styles.typeBadgeRow}>
                    <span className={styles.typeBadge}>{announcement.HighlightType ?? 'General'}</span>
                    <span className={`${styles.priorityBadge} ${
                      announcement.Priority === 'Critical' ? styles.priorityCritical :
                      announcement.Priority === 'High'     ? styles.priorityHigh :
                      announcement.Priority === 'Medium'   ? styles.priorityMedium :
                      styles.priorityLow
                    }`}>
                      {announcement.Priority}
                    </span>
                  </div>
                  <h1 className={styles.tabbedAnnouncementTitle}>{announcement.Title ?? 'Untitled'}</h1>
                  {announcement.Author && (
                    <span className={styles.authorLabel}>
                      <Icon iconName="Contact" /> {announcement.Author.Title}
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* Edit / locked controls */}
            <div className={styles.articleActions}>
              {canEdit ? (
                <PrimaryButton
                  text="Edit"
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => onEdit(announcement)}
                  className={styles.editButton}
                />
              ) : (
                <div className={styles.lockedNote}>
                  <Icon iconName="Lock" /> <span>Only editors and the creator can edit</span>
                </div>
              )}
            </div>

            {/* Body — rich text rendered as-is from SharePoint */}
            {announcement.Body && (
              <div
                className={styles.bodyBox}
                dangerouslySetInnerHTML={{ __html: announcement.Body }}
              />
            )}

            {/* Attachments */}
            {announcement.Attachments && announcement.Attachments.length > 0 && (
              <div className={styles.infoCard}>
                <div className={styles.cardHeader}>
                  <Icon iconName="Attach" className={styles.cardIcon} />
                  <h3>Attachments ({announcement.Attachments.length})</h3>
                </div>
                <div className={styles.cardContent}>
                  <div className={styles.attachmentList}>
                    {announcement.Attachments.map((att, idx) => (
                      <a
                        key={idx}
                        href={att.ServerRelativeUrl}
                        target="_blank"
                        rel="noreferrer"
                        className={styles.attachmentCard}
                      >
                        <div className={styles.attachmentIconWrapper}>
                          <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentTypeIcon} />
                        </div>
                        <div className={styles.attachmentDetails}>
                          <span className={styles.attachmentName}>{att.FileName}</span>
                          <span className={styles.attachmentAction}>Click to open</span>
                        </div>
                        <Icon iconName="OpenInNewWindow" className={styles.attachmentOpenIcon} />
                      </a>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* Status footer */}
            <div className={styles.statusContainer}>
              <span className={`${styles.statusBadge} ${
                announcement.Status === 'Published' ? styles.statusActive : styles.statusDraft
              }`}>
                {announcement.Status ?? 'Draft'}
              </span>
            </div>

          </div>
        </div>
      </div>
    </div>,
    document.body
  );
};

export default TabbedAnnouncementDetailsPanel;