import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./AnnouncementsCarousel.module.scss";
import { IAnnouncementsCarouselProps } from "./IAnnouncementsCarouselProps";
import { AnnouncementService } from "../services/AnnouncementService";
import { IAnnouncement } from "../models/IAnnouncement";
import AnnouncementDetailsPanel from "./AnnouncementDetailsPanel";
import ViewAllPanel from "./ViewAllPanel";
import AddEditAnnouncement from "./AddEditAnnouncement";

import SwiperCore, { Autoplay, Pagination } from "swiper";
import { Swiper, SwiperSlide } from "swiper/react";
import "swiper/css";
import "swiper/css/pagination";

SwiperCore.use([Autoplay, Pagination]);

const stripHtml = (html: string): string =>
  html.replace(/<[^>]+>/g, "").trim();

const AnnouncementsCarousel = (props: IAnnouncementsCarouselProps): JSX.Element => {
  const [announcements, setAnnouncements] = useState<IAnnouncement[]>([]);
  const [selectedAnnouncement, setSelectedAnnouncement] = useState<IAnnouncement | undefined>(undefined);
  const [isDetailsOpen, setIsDetailsOpen] = useState(false);
  const [isViewAllOpen, setIsViewAllOpen] = useState(false);
  const [isAddEditOpen, setIsAddEditOpen] = useState(false);
  const [addEditMode, setAddEditMode] = useState<"add" | "edit">("add");
  const [announcementToEdit, setAnnouncementToEdit] = useState<IAnnouncement | undefined>(undefined);

  const loadAnnouncements = async (): Promise<void> => {
    try {
      const service = new AnnouncementService(props.context, props.sourceList);
      const data = await service.getAnnouncements(props.currentUserLogin);
      setAnnouncements(data);
    } catch (err) {
      console.error("Failed to load announcements", err);
    }
  };

  useEffect(() => {
    if (!props.sourceList) return;
    loadAnnouncements().catch(console.error);
  }, [props.sourceList, props.context, props.currentUserLogin]);

  const limit = Math.min(props.carouselLimit || 3, 5);

  const carouselAnnouncements = announcements
    .filter(a => a.Status?.trim() === "Active")
    .slice(0, limit);

  const openDetailsFromCarousel = (announcement: IAnnouncement): void => {
    setSelectedAnnouncement(announcement);
    setIsDetailsOpen(true);
  };

  const openDetailsFromViewAll = (announcement: IAnnouncement): void => {
    setSelectedAnnouncement(announcement);
    setIsDetailsOpen(true);
  };

  const closeDetails = (): void => {
    setIsDetailsOpen(false);
    setSelectedAnnouncement(undefined);
  };

  const openAddAnnouncement = (): void => {
    setAddEditMode("add");
    setAnnouncementToEdit(undefined);
    setIsAddEditOpen(true);
  };

  const openEditAnnouncement = (announcement: IAnnouncement): void => {
    setIsDetailsOpen(false);
    setAddEditMode("edit");
    setAnnouncementToEdit(announcement);
    setIsAddEditOpen(true);
  };

  const closeAddEditPanel = (): void => {
    setIsAddEditOpen(false);
    if (announcementToEdit) {
      setSelectedAnnouncement(announcementToEdit);
      setIsDetailsOpen(true);
    }
    setAnnouncementToEdit(undefined);
  };

  const handleSaveAnnouncement = async (
    announcementData: Partial<IAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> => {
    const service = new AnnouncementService(props.context, props.sourceList);
    try {
      if (addEditMode === "add") {
        await service.addAnnouncement(announcementData, bannerFile, attachments);
      } else if (addEditMode === "edit" && announcementToEdit) {
        await service.updateAnnouncement(announcementToEdit.Id, announcementData, bannerFile, attachments, deletedAttachmentNames);
      }
      setIsAddEditOpen(false);
      setAnnouncementToEdit(undefined);
      setSelectedAnnouncement(undefined);
      setIsDetailsOpen(false);
      setIsViewAllOpen(false);
      await loadAnnouncements();
    } catch (err) {
      console.error("Failed to save announcement:", err);
      throw err;
    }
  };

  const renderPanels = (): JSX.Element => (
    <>
      <ViewAllPanel
        announcements={announcements}
        isOpen={isViewAllOpen}
        onDismiss={() => setIsViewAllOpen(false)}
        onSelectAnnouncement={openDetailsFromViewAll}
        onAddAnnouncement={openAddAnnouncement}
      />
      <AnnouncementDetailsPanel
        announcement={selectedAnnouncement}
        isOpen={isDetailsOpen}
        onDismiss={closeDetails}
        onEdit={openEditAnnouncement}
        currentUserId={props.currentUserId}
      />
      <AddEditAnnouncement
        isOpen={isAddEditOpen}
        mode={addEditMode}
        announcement={announcementToEdit}
        context={props.context}
        onDismiss={closeAddEditPanel}
        onSave={handleSaveAnnouncement}
      />
    </>
  );

  if (!carouselAnnouncements.length) {
    return (
      <div className={styles.carouselWrapper}>
        <div className={styles.noAnnouncements}>
          <div className={styles.noAnnouncementsIcon}>📢</div>
          {!props.sourceList ? (
            <p>Please select a list in the web part settings to get started.</p>
          ) : (
            <>
              <p>No active announcements</p>
              <div className={styles.noAnnouncementsActions}>
                <button className={styles.noAnnouncementsViewAll} onClick={() => setIsViewAllOpen(true)}>
                  View All Announcements
                </button>
                <button className={styles.noAnnouncementsAdd} onClick={openAddAnnouncement}>
                  + Add Announcement
                </button>
              </div>
            </>
          )}
        </div>
        {props.sourceList && renderPanels()}
      </div>
    );
  }

  return (
    <div className={styles.carouselWrapper}>
      <Swiper
        watchOverflow={false}
        spaceBetween={0}
        slidesPerView={1}
        loop={carouselAnnouncements.length > 1}
        autoplay={{ delay: 6000, disableOnInteraction: false, pauseOnMouseEnter: true }}
        pagination={{
          clickable: true,
          el: `.${styles.swiperDots}`,
          bulletClass: "carousel-dot",
          bulletActiveClass: "carousel-dot-active",
        }}
      >
        {carouselAnnouncements.map((announcement) => {
          const hasBanner = !!announcement.BannerImageUrl;
          const bodyText = stripHtml(announcement.Body ?? "");
          const readingTime = Math.max(1, Math.ceil((bodyText.length || 1) / 200));

          return (
            <SwiperSlide key={announcement.Id}>
              <div className={`${styles.card} ${hasBanner ? styles.cardHasBanner : ""}`}>

                {/* ── Banner image — actual <img> so Swiper loop cloning works ── */}
                {hasBanner && (
                  <img
                    src={announcement.BannerImageUrl}
                    alt=""
                    aria-hidden="true"
                    className={styles.bannerBg}
                  />
                )}

                {/* ── Gradient overlay — always rendered ── */}
                <div className={styles.bannerGradient} />

                {/* ── Card content — all left-aligned ── */}
                <div className={styles.cardInner}>

                  {/* Zone 1 — meta row glued to top */}
                  <div className={styles.topMeta}>
                    <span className={styles.typeLabel}>{announcement.AnnouncementType}</span>
                    <span className={styles.metaSep}>·</span>
                    <span className={styles.metaItem}>
                      {announcement.PublishDate.toLocaleDateString("en-US", {
                        month: "short",
                        day: "numeric",
                        year: "numeric",
                      })}
                    </span>
                    <span className={styles.metaSep}>·</span>
                    <span className={styles.metaItem}>{readingTime} min read</span>
                  </div>

                  {/* Zone 2 — title + teaser centred in the middle */}
                  <div className={styles.cardMiddle}>
                    <h2 className={styles.title}>{announcement.Title}</h2>
                    {bodyText && (
                      <p className={styles.teaser}>{bodyText}</p>
                    )}
                  </div>

                  {/* Zone 3 — actions + pagination pinned to bottom */}
                  <div className={styles.cardBottom}>
                    <div className={styles.cardActions}>
                      <button
                        className={styles.readMoreBtn}
                        onClick={() => openDetailsFromCarousel(announcement)}
                      >
                        Read More →
                      </button>
                      <button
                        className={styles.viewAllBtnCard}
                        onClick={() => setIsViewAllOpen(true)}
                      >
                        View All
                      </button>
                    </div>
                    <div className={styles.swiperDots} />
                  </div>

                </div>

              </div>
            </SwiperSlide>
          );
        })}
      </Swiper>

      {renderPanels()}
    </div>
  );
};

export default AnnouncementsCarousel;