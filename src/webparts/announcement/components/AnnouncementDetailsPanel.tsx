import * as React from 'react';
import { PrimaryButton, Icon } from '@fluentui/react';
import { IAnnouncement } from '../models/IAnnouncement';
import styles from './AnnouncementDetailsPanel.module.scss';

interface AnnouncementDetailsPanelProps {
  announcement?: IAnnouncement;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (announcement: IAnnouncement) => void;
  currentUserId: number;
}

const AnnouncementDetailsPanel: React.FC<AnnouncementDetailsPanelProps> = ({
  announcement, isOpen, onDismiss, onEdit, currentUserId,
}) => {
  if (!isOpen || !announcement) return null; // eslint-disable-line @rushstack/no-new-null

  const isExpired = announcement.Status === 'Expired';
  const isAuthor = announcement.Author?.Id === currentUserId;
  const canEdit = !isExpired && isAuthor;

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

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Hero Banner — title overlaid on image (article style) ── */}
        {hasBanner ? (
          <div className={styles.bannerContainer}>
            <img
              src={announcement.BannerImageUrl}
              alt={announcement.Title ?? 'Announcement Banner'}
              className={styles.bannerImage}
            />

            {/* gradient + title/meta overlaid on banner */}
            <div className={styles.bannerOverlayContent}>
              <div className={styles.bannerTypeBadgeRow}>
                <span className={styles.bannerTypeBadge}>{announcement.AnnouncementType ?? 'General'}</span>
                <span className={styles.bannerTypeBadge}>{announcement.Priority}</span>
              </div>
              <h1 className={styles.bannerTitle}>{announcement.Title ?? 'Untitled'}</h1>
              {announcement.Author && (
                <span className={styles.bannerAuthor}>
                  <Icon iconName="Contact" /> {announcement.Author.Title}
                </span>
              )}
              <span className={styles.bannerPublishDate}>
                {announcement.PublishDate.toLocaleDateString('en-US', {
                  weekday: 'short', year: 'numeric', month: 'short', day: 'numeric',
                })}
              </span>
            </div>

            {isExpired && (
              <div className={styles.expiredOverlay}>
                <span>Expired</span>
              </div>
            )}

            <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
              <Icon iconName="Cancel" />
            </button>
          </div>
        ) : (
          // No banner: close button anchored to top-right of modal
          <button
            className={`${styles.closeBtn} ${styles.closeBtnNoImage}`}
            style={{ position: 'absolute', top: 14, right: 14, zIndex: 10 }}
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
                    <span className={styles.typeBadge}>{announcement.AnnouncementType ?? 'General'}</span>
                    <span className={`${styles.priorityBadge} ${
                      announcement.Priority === 'Critical' ? styles.priorityCritical :
                      announcement.Priority === 'High'     ? styles.priorityHigh :
                      announcement.Priority === 'Medium'   ? styles.priorityMedium :
                      styles.priorityLow
                    }`}>
                      {announcement.Priority}
                    </span>
                  </div>
                  <h1 className={styles.announcementTitle}>{announcement.Title ?? 'Untitled'}</h1>
                  {announcement.Author && (
                    <span className={styles.authorLabel}>
                      <Icon iconName="Contact" /> {announcement.Author.Title}
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* Edit / locked controls */}
            {(canEdit || (!canEdit && isExpired) || (!canEdit && !isExpired && !isAuthor)) && (
              <div className={styles.articleActions}>
                {canEdit && (
                  <PrimaryButton
                    text="Edit"
                    iconProps={{ iconName: 'Edit' }}
                    onClick={() => onEdit(announcement)}
                    className={styles.editButton}
                  />
                )}
                {!canEdit && isExpired && (
                  <div className={styles.lockedNote}>
                    <Icon iconName="Lock" /> <span>Edit via SharePoint list</span>
                  </div>
                )}
                {!canEdit && !isExpired && !isAuthor && (
                  <div className={styles.lockedNote}>
                    <Icon iconName="Lock" /> <span>Only the creator can edit</span>
                  </div>
                )}
              </div>
            )}

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
                announcement.Status === 'Active'    ? styles.statusActive :
                announcement.Status === 'Scheduled' ? styles.statusScheduled :
                announcement.Status === 'Expired'   ? styles.statusExpired :
                styles.statusDraft
              }`}>
                {announcement.Status ?? 'Draft'}
              </span>
            </div>

          </div>
        </div>
      </div>
    </div>
  );
};

export default AnnouncementDetailsPanel;