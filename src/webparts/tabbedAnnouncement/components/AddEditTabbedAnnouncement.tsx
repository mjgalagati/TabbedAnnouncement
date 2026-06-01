import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import * as ReactDOM from 'react-dom';
import {
  PrimaryButton, DefaultButton,
  TextField, Dropdown, IDropdownOption, Icon,
  Spinner, SpinnerSize, MessageBar, MessageBarType,
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { ITabbedAnnouncement, ITabbedAnnouncementAttachment, ITabbedAnnouncementAudience } from '../models/ITabbedAnnouncement';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './AddEditTabbedAnnouncement.module.scss';

interface AddEditTabbedAnnouncementProps {
  isOpen: boolean;
  mode: 'add' | 'edit';
  announcement?: ITabbedAnnouncement;
  context: WebPartContext;
  highlightTypeOptions: string[];
  siteOptions: string[];
  onDismiss: () => void;
  onSave: (
    data: Partial<ITabbedAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ) => Promise<void>;
}

const PRIORITY_OPTIONS: IDropdownOption[] = [
  { key: 'Critical', text: 'Critical' },
  { key: 'High',     text: 'High' },
  { key: 'Medium',   text: 'Medium' },
  { key: 'Low',      text: 'Low' },
];

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Draft',     text: 'Draft' },
  { key: 'Published', text: 'Published' },
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

const AddEditTabbedAnnouncement: React.FC<AddEditTabbedAnnouncementProps> = ({
  isOpen, mode, announcement, context, highlightTypeOptions, siteOptions, onDismiss, onSave,
}) => {
  const bodyEditorRef = useRef<HTMLDivElement>(null);

  const [title, setTitle] = useState('');
  const [highlightType, setHighlightType] = useState('');
  const [site, setSite] = useState('All');
  const [priority, setPriority] = useState<'Critical' | 'High' | 'Medium' | 'Low'>('Medium');
  const [status, setStatus] = useState<'Draft' | 'Published'>('Draft');
  const [targetAudienceType, setTargetAudienceType] = useState<'All' | 'Specific' | 'Except'>('All');
  const [targetAudience, setTargetAudience] = useState<ITabbedAnnouncementAudience[]>([]);
  const [bannerFile, setBannerFile] = useState<File | undefined>(undefined);
  const [bannerPreview, setBannerPreview] = useState<string | undefined>(undefined);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<ITabbedAnnouncementAttachment[]>([]);
  const [deletedAttachmentNames, setDeletedAttachmentNames] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [saveError, setSaveError] = useState<string | undefined>(undefined);

  const highlightTypeDropdownOptions: IDropdownOption[] = highlightTypeOptions.map(t => ({ key: t, text: t }));
  const siteDropdownOptions: IDropdownOption[] = [
    { key: 'All', text: 'All (appears in every tab)' },
    ...siteOptions.map(s => ({ key: s, text: s })),
  ];

  useEffect(() => {
    if (isOpen) {
      if (mode === 'edit' && announcement) {
        setTitle(announcement.Title ?? '');
        setHighlightType(announcement.HighlightType ?? '');
        setSite(announcement.Site ?? 'All');
        setPriority(announcement.Priority ?? 'Medium');
        setStatus((announcement.Status === 'Draft' || announcement.Status === 'Published') ? announcement.Status : 'Draft');
        setTargetAudienceType(announcement.TargetAudienceType ?? 'All');
        setTargetAudience(announcement.TargetAudience ?? []);
        setBannerPreview(announcement.BannerImageUrl ?? undefined);
        setExistingAttachments(announcement.Attachments ?? []);
        if (bodyEditorRef.current) bodyEditorRef.current.innerHTML = announcement.Body ?? '';
      } else {
        setTitle('');
        setHighlightType(highlightTypeOptions[0] ?? '');
        setSite('All');
        setPriority('Medium');
        setStatus('Draft');
        setTargetAudienceType('All');
        setTargetAudience([]);
        setBannerPreview(undefined);
        setExistingAttachments([]);
        if (bodyEditorRef.current) bodyEditorRef.current.innerHTML = '';
      }
      setBannerFile(undefined);
      setNewAttachments([]);
      setDeletedAttachmentNames([]);
      setErrors({});
      setSaveError(undefined);
    }
  }, [isOpen, mode, announcement]);

  if (!isOpen) return null; // eslint-disable-line @rushstack/no-new-null

  const execFormat = (command: string): void => {
    document.execCommand(command, false, undefined);
  };

  const validate = (): boolean => {
    const newErrors: { [key: string]: string } = {};
    if (!title.trim()) newErrors.title = 'Title is required';
    if (!highlightType) newErrors.highlightType = 'Highlight type is required';
    if (!site) newErrors.site = 'Site is required';
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
      const data: Partial<ITabbedAnnouncement> = {
        Title: title,
        Body: bodyEditorRef.current?.innerHTML ?? '',
        HighlightType: highlightType,
        Site: site,
        Priority: priority,
        Status: status,
        TargetAudienceType: targetAudienceType,
        TargetAudience: targetAudience,
        BannerImageUrl: bannerPreview,
      };
      await onSave(data, bannerFile, newAttachments, deletedAttachmentNames);
    } catch (err) {
      console.error('Save error:', err);
      setSaveError('Failed to save highlight. Please try again.');
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

  const modal = (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <h2 className={styles.panelTitle}>
            <Icon iconName={mode === 'add' ? 'Add' : 'Edit'} className={styles.titleIcon} />
            {mode === 'add' ? 'New Highlight' : 'Edit Highlight'}
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
              placeholder="Enter highlight title"
            />
            <Dropdown
              label="Highlight Type"
              required
              selectedKey={highlightType}
              options={highlightTypeDropdownOptions}
              onChange={(_, o) => setHighlightType(o?.key as string ?? '')}
              errorMessage={errors.highlightType}
              disabled={highlightTypeDropdownOptions.length === 0}
              placeholder={highlightTypeDropdownOptions.length === 0 ? 'Configure Highlight Types in web part settings' : 'Select highlight type'}
            />
            <Dropdown
              label="Site"
              required
              selectedKey={site}
              options={siteDropdownOptions}
              onChange={(_, o) => setSite(o?.key as string ?? 'All')}
              errorMessage={errors.site}
              disabled={siteDropdownOptions.length === 0}
              placeholder="Select site"
            />
          </div>

          {/* Body */}
          <div className={styles.formSection}>
            <div className={styles.bodyField}>
              <label className={styles.bodyLabel}>Body</label>
              <div className={styles.bodyToolbar}>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Bold"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('bold'); }}
                >
                  <b>B</b>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Italic"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('italic'); }}
                >
                  <i>I</i>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Underline"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('underline'); }}
                >
                  <u>U</u>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Bullet list"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('insertUnorderedList'); }}
                >
                  <Icon iconName="BulletedList" />
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Numbered list"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('insertOrderedList'); }}
                >
                  <Icon iconName="NumberedList" />
                </button>
              </div>
              <div
                ref={bodyEditorRef}
                className={styles.bodyEditor}
                contentEditable={true}
                data-placeholder="Write your message here…"
                suppressContentEditableWarning={true}
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
          </div>

          {/* Scheduling */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Calendar" className={styles.sectionIcon} />
              <h3>Scheduling</h3>
            </div>
            <Dropdown
              label="Status"
              selectedKey={status}
              options={STATUS_OPTIONS}
              onChange={(_, o) => setStatus(o?.key as 'Draft' | 'Published' ?? 'Draft')}
            />
          </div>

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
                Only the selected people below will see this highlight.
              </span>
            )}
            {targetAudienceType === 'Except' && (
              <span className={styles.fieldHint}>
                Everyone will see this highlight <strong>except</strong> the selected people below.
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
                    const people: ITabbedAnnouncementAudience[] = (items || []).map((item: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
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
            text={isSaving ? 'Saving...' : mode === 'add' ? 'Add Highlight' : 'Save Changes'}
            onClick={handleSave}
            disabled={isSaving}
          >
            {isSaving && <Spinner size={SpinnerSize.small} />}
          </PrimaryButton>
        </div>

      </div>
    </div>
  );

  return ReactDOM.createPortal(modal, document.body);
};

export default AddEditTabbedAnnouncement;
