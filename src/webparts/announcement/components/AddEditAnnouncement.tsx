import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  PrimaryButton, DefaultButton,
  TextField, Dropdown, IDropdownOption, Icon,
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  DatePicker, IconButton,
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { IAnnouncement, IAnnouncementAttachment, IAnnouncementAudience } from '../models/IAnnouncement';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './AddEditAnnouncement.module.scss';

interface AddEditAnnouncementProps {
  isOpen: boolean;
  mode: 'add' | 'edit';
  announcement?: IAnnouncement;
  context: WebPartContext;
  onDismiss: () => void;
  onSave: (
    data: Partial<IAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ) => Promise<void>;
}

const ANNOUNCEMENT_TYPES: IDropdownOption[] = [
  { key: 'Company News',    text: 'Company News' },
  { key: 'Department',      text: 'Department' },
  { key: 'Policy/HR',       text: 'Policy/HR' },
  { key: 'Product/Project', text: 'Product/Project' },
  { key: 'General',         text: 'General' },
];

const PRIORITY_OPTIONS: IDropdownOption[] = [
  { key: 'Critical', text: 'Critical' },
  { key: 'High',     text: 'High' },
  { key: 'Medium',   text: 'Medium' },
  { key: 'Low',      text: 'Low' },
];

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Draft',     text: 'Draft' },
  { key: 'Scheduled', text: 'Scheduled' },
];

const AUDIENCE_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'All',      text: 'All' },
  { key: 'Specific', text: 'Specific' },
  { key: 'Except',   text: 'Except' },
];

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

const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
};

const AddEditAnnouncement: React.FC<AddEditAnnouncementProps> = ({
  isOpen, mode, announcement, context, onDismiss, onSave,
}) => {
  const [title, setTitle] = useState('');
  const [body, setBody] = useState('');
  const [announcementType, setAnnouncementType] = useState('General');
  const [publishDate, setPublishDate] = useState<Date | undefined>(new Date());
  const [expiryDate, setExpiryDate] = useState<Date | undefined>(undefined);
  const [priority, setPriority] = useState<'Critical' | 'High' | 'Medium' | 'Low'>('Medium');
  const [status, setStatus] = useState<'Draft' | 'Scheduled'>('Draft');
  const [targetAudienceType, setTargetAudienceType] = useState<'All' | 'Specific' | 'Except'>('All');
  const [targetAudience, setTargetAudience] = useState<IAnnouncementAudience[]>([]);
  const [bannerFile, setBannerFile] = useState<File | undefined>(undefined);
  const [bannerPreview, setBannerPreview] = useState<string | undefined>(undefined);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<IAnnouncementAttachment[]>([]);
  const [deletedAttachmentNames, setDeletedAttachmentNames] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [saveError, setSaveError] = useState<string | undefined>(undefined);

  const isCritical = priority === 'Critical';

  useEffect(() => {
    if (isOpen) {
      if (mode === 'edit' && announcement) {
        setTitle(announcement.Title ?? '');
        setBody(announcement.Body ?? '');
        setAnnouncementType(announcement.AnnouncementType ?? 'General');
        setPublishDate(announcement.PublishDate ?? new Date());
        setExpiryDate(announcement.ExpiryDate ?? undefined);
        setPriority(announcement.Priority ?? 'Medium');
        setStatus((announcement.Status === 'Draft' || announcement.Status === 'Scheduled') ? announcement.Status : 'Draft');
        setTargetAudienceType(announcement.TargetAudienceType ?? 'All');
        setTargetAudience(announcement.TargetAudience ?? []);
        setBannerPreview(announcement.BannerImageUrl ?? undefined);
        setExistingAttachments(announcement.Attachments ?? []);
      } else {
        setTitle('');
        setBody('');
        setAnnouncementType('General');
        setPublishDate(new Date());
        setExpiryDate(undefined);
        setPriority('Medium');
        setStatus('Draft');
        setTargetAudienceType('All');
        setTargetAudience([]);
        setBannerPreview(undefined);
        setExistingAttachments([]);
      }
      setBannerFile(undefined);
      setNewAttachments([]);
      setDeletedAttachmentNames([]);
      setErrors({});
      setSaveError(undefined);
    }
  }, [isOpen, mode, announcement]);

  if (!isOpen) return null; // eslint-disable-line @rushstack/no-new-null

  const validate = (): boolean => {
    const newErrors: { [key: string]: string } = {};
    if (!title.trim()) newErrors.title = 'Title is required';
    if (!announcementType) newErrors.announcementType = 'Announcement type is required';
    if (!isCritical && !publishDate) newErrors.publishDate = 'Publish date is required';
    if (isCritical && !expiryDate) newErrors.expiryDate = 'Expiry date is required for Critical announcements';
    if (expiryDate && publishDate && expiryDate <= publishDate) {
      newErrors.expiryDate = 'Expiry date must be after publish date';
    }
    if (targetAudienceType !== 'All' && targetAudience.length === 0) {
      newErrors.targetAudience = 'Please select at least one person for the target audience';
    }
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setIsSaving(true);
    setSaveError(undefined);
    try {
      const data: Partial<IAnnouncement> = {
        Title: title,
        Body: body,
        AnnouncementType: announcementType,
        PublishDate: isCritical ? new Date() : publishDate,
        ExpiryDate: expiryDate,
        Priority: priority,
        Status: isCritical ? 'Active' : status,
        TargetAudienceType: targetAudienceType,
        TargetAudience: targetAudience,
        BannerImageUrl: bannerPreview,
      };
      await onSave(data, bannerFile, newAttachments, deletedAttachmentNames);
    } catch (err) {
      console.error('Save error:', err);
      setSaveError('Failed to save announcement. Please try again.');
    } finally {
      setIsSaving(false);
    }
  };

  const handleBannerChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBannerFile(file);
    const reader = new FileReader();
    reader.onload = (ev) => setBannerPreview(ev.target?.result as string);
    reader.readAsDataURL(file);
  };

  const handleRemoveBanner = (): void => {
    setBannerFile(undefined);
    setBannerPreview(undefined);
  };

  const handleAttachmentAdd = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const files = Array.from(e.target.files ?? []);
    if (!files.length) return;
    setNewAttachments(prev => [...prev, ...files]);
  };

  const handleRemoveNewAttachment = (index: number): void => {
    setNewAttachments(prev => prev.filter((_, i) => i !== index));
  };

  const handleRemoveExistingAttachment = (fileName: string): void => {
    setExistingAttachments(prev => prev.filter(a => a.FileName !== fileName));
    setDeletedAttachmentNames(prev => [...prev, fileName]);
  };

  const totalAttachments = existingAttachments.length + newAttachments.length;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <h2 className={styles.panelTitle}>
            <Icon iconName={mode === 'add' ? 'Add' : 'Edit'} className={styles.titleIcon} />
            {mode === 'add' ? 'New Announcement' : 'Edit Announcement'}
          </h2>
          <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
            <Icon iconName="Cancel" />
          </button>
        </div>

        {/* ── Scrollable Form Body ── */}
        <div className={styles.panelContent}>
          {saveError && (
            <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar}>
              {saveError}
            </MessageBar>
          )}

          {/* Basic Info */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Info" className={styles.sectionIcon} />
              <h3>Basic Information</h3>
            </div>
            <TextField
              label="Title"
              required
              value={title}
              onChange={(_, v) => setTitle(v ?? '')}
              errorMessage={errors.title}
              placeholder="Enter announcement title"
            />
            <Dropdown
              label="Announcement Type"
              required
              selectedKey={announcementType}
              options={ANNOUNCEMENT_TYPES}
              onChange={(_, o) => setAnnouncementType(o?.key as string ?? 'General')}
              errorMessage={errors.announcementType}
            />
          </div>

          {/* Body — Rich Text */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="AlignLeft" className={styles.sectionIcon} />
              <h3>Body</h3>
            </div>
            <div className={styles.richTextWrapper}>
              <RichText
                value={body}
                onChange={(newValue) => { setBody(newValue ?? ''); return newValue ?? ''; }}
                isEditMode={true}
              />
            </div>
          </div>

          {/* Banner Image */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Photo2" className={styles.sectionIcon} />
              <h3>Banner Image</h3>
            </div>
            <div className={styles.bannerUpload}>
              {bannerPreview ? (
                <div className={styles.bannerPreviewWrapper}>
                  <img src={bannerPreview} alt="Banner preview" className={styles.bannerPreview} />
                  <DefaultButton
                    text="Remove Banner"
                    iconProps={{ iconName: 'Delete' }}
                    onClick={handleRemoveBanner}
                    className={styles.removeBtn}
                  />
                </div>
              ) : (
                <label className={styles.uploadArea} htmlFor="banner-upload">
                  <Icon iconName="Photo2" className={styles.uploadIcon} />
                  <span className={styles.uploadLabel}>Click to upload banner image</span>
                  <span className={styles.uploadHint}>JPG, PNG, GIF, WEBP supported</span>
                  <input
                    id="banner-upload"
                    type="file"
                    accept="image/*"
                    className={styles.fileInput}
                    onChange={handleBannerChange}
                  />
                </label>
              )}
            </div>
          </div>

          {/* Priority */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Flag" className={styles.sectionIcon} />
              <h3>Priority</h3>
            </div>
            <Dropdown
              label="Priority Level"
              selectedKey={priority}
              options={PRIORITY_OPTIONS}
              onChange={(_, o) => setPriority(o?.key as 'Critical' | 'High' | 'Medium' | 'Low' ?? 'Medium')}
            />
            {isCritical && (
              <MessageBar messageBarType={MessageBarType.severeWarning} className={styles.criticalWarning}>
                <strong>Critical Priority — Post Now!</strong> This announcement will be published
                immediately upon saving, bypassing Draft and Scheduled status. It will supersede all
                other announcements in the carousel regardless of publish date. An expiry date is required.
              </MessageBar>
            )}
          </div>

          {/* Scheduling — hidden for Critical */}
          {!isCritical && (
            <div className={styles.formSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Calendar" className={styles.sectionIcon} />
                <h3>Scheduling</h3>
              </div>
              <Dropdown
                label="Status"
                selectedKey={status}
                options={STATUS_OPTIONS}
                onChange={(_, o) => setStatus(o?.key as 'Draft' | 'Scheduled' ?? 'Draft')}
              />
              <div className={styles.dateRow}>
                <div className={styles.dateField}>
                  <DatePicker
                    label="Publish Date *"
                    value={publishDate}
                    onSelectDate={(d) => setPublishDate(d ?? undefined)}
                    formatDate={(d) => d?.toLocaleDateString() ?? ''}
                    placeholder="Select publish date"
                  />
                  {errors.publishDate && <span className={styles.error}>{errors.publishDate}</span>}
                </div>
                <div className={styles.dateField}>
                  <DatePicker
                    label="Expiry Date"
                    value={expiryDate}
                    onSelectDate={(d) => setExpiryDate(d ?? undefined)}
                    formatDate={(d) => d?.toLocaleDateString() ?? ''}
                    placeholder="Select expiry date (optional)"
                    minDate={publishDate}
                  />
                  {errors.expiryDate && <span className={styles.error}>{errors.expiryDate}</span>}
                  {expiryDate && (
                    <IconButton
                      iconProps={{ iconName: 'Clear' }}
                      onClick={() => setExpiryDate(undefined)}
                      title="Clear expiry date"
                    />
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Expiry only — Critical */}
          {isCritical && (
            <div className={styles.formSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Calendar" className={styles.sectionIcon} />
                <h3>Expiry Date</h3>
              </div>
              <DatePicker
                label="Expiry Date *"
                value={expiryDate}
                onSelectDate={(d) => setExpiryDate(d ?? undefined)}
                formatDate={(d) => d?.toLocaleDateString() ?? ''}
                placeholder="Select expiry date"
                minDate={new Date()}
              />
              {errors.expiryDate && <span className={styles.error}>{errors.expiryDate}</span>}
            </div>
          )}

          {/* Target Audience */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="People" className={styles.sectionIcon} />
              <h3>Target Audience</h3>
            </div>
            <Dropdown
              label="Audience Type"
              selectedKey={targetAudienceType}
              options={AUDIENCE_TYPE_OPTIONS}
              onChange={(_, o) => {
                setTargetAudienceType(o?.key as 'All' | 'Specific' | 'Except' ?? 'All');
                setTargetAudience([]);
              }}
            />
            {targetAudienceType === 'Specific' && (
              <span className={styles.fieldHint}>
                Only the selected people below will see this announcement.
              </span>
            )}
            {targetAudienceType === 'Except' && (
              <span className={styles.fieldHint}>
                Everyone will see this announcement <strong>except</strong> the selected people below.
              </span>
            )}
            {targetAudienceType !== 'All' && (
              <>
                <PeoplePicker
                  context={context as any} // eslint-disable-line @typescript-eslint/no-explicit-any
                  titleText="Select People"
                  personSelectionLimit={50}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  defaultSelectedUsers={targetAudience.map(p => p.Title)}
                  onChange={(items) => {
                    const people: IAnnouncementAudience[] = (items || []).map((item: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
                      Id: parseInt(item.id, 10),
                      Title: item.text,
                    }));
                    setTargetAudience(people);
                  }}
                />
                {errors.targetAudience && <span className={styles.error}>{errors.targetAudience}</span>}
              </>
            )}
          </div>

          {/* Attachments */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Attach" className={styles.sectionIcon} />
              <h3>
                Attachments
                {totalAttachments > 0 && (
                  <span className={styles.attachmentCount}>{totalAttachments}</span>
                )}
              </h3>
            </div>

            {existingAttachments.length > 0 && (
              <div className={styles.attachmentList}>
                {existingAttachments.map((att) => (
                  <div key={att.FileName} className={styles.attachmentItem}>
                    <div className={styles.attachmentInfo}>
                      <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentIcon} />
                      <span className={styles.attachmentName}>{att.FileName}</span>
                    </div>
                    <div className={styles.attachmentActions}>
                      <a href={att.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.attachmentDownload}>
                        <Icon iconName="Download" />
                      </a>
                      <button className={styles.attachmentRemove} onClick={() => handleRemoveExistingAttachment(att.FileName)}>
                        <Icon iconName="Delete" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {newAttachments.length > 0 && (
              <div className={styles.attachmentList}>
                {newAttachments.map((file, idx) => (
                  <div key={idx} className={`${styles.attachmentItem} ${styles.attachmentNew}`}>
                    <div className={styles.attachmentInfo}>
                      <Icon iconName={getFileIcon(file.name)} className={styles.attachmentIcon} />
                      <span className={styles.attachmentName}>{file.name}</span>
                      <span className={styles.attachmentSize}>{formatFileSize(file.size)}</span>
                    </div>
                    <div className={styles.attachmentActions}>
                      <button className={styles.attachmentRemove} onClick={() => handleRemoveNewAttachment(idx)}>
                        <Icon iconName="Delete" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            <label className={styles.attachmentUploadArea} htmlFor="attachment-upload">
              <Icon iconName="Attach" />
              <span>Click to add attachments</span>
              <input
                id="attachment-upload"
                type="file"
                multiple
                className={styles.fileInput}
                onChange={handleAttachmentAdd}
              />
            </label>
          </div>
        </div>

        {/* ── Sticky Action Buttons ── */}
        <div className={styles.actionButtons}>
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSaving} />
          <PrimaryButton
            text={isSaving ? 'Saving...' : mode === 'add' ? 'Add Announcement' : 'Save Changes'}
            onClick={handleSave}
            disabled={isSaving}
          >
            {isSaving && <Spinner size={SpinnerSize.small} />}
          </PrimaryButton>
        </div>

      </div>
    </div>
  );
};

export default AddEditAnnouncement;