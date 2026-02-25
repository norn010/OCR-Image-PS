import React, { useEffect, useRef, useState } from "react";
import "./index.css";

// ใช้ path แบบ relative แล้วให้ Vite proxy ไปที่ Flask (ดูไฟล์ vite.config.ts)
const API_BASE = "";

import type { OcrStatus, ExcelSheet } from "./types";
import { UploadPanel } from "./components/UploadPanel";
import { ImagePreviewPanel } from "./components/ImagePreviewPanel";
import { OcrResultPanel } from "./components/OcrResultPanel";
import { Lightbox } from "./components/Lightbox";

const App: React.FC = () => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedFileName, setSelectedFileName] = useState<string>("ยังไม่ได้เลือกไฟล์");
  const [imagePreviewUrl, setImagePreviewUrl] = useState<string | null>(null);

  const [jobId, setJobId] = useState<string>("");
  const [status, setStatus] = useState<OcrStatus | null>(null);
  const [isStarting, setIsStarting] = useState(false);

  const [resultId, setResultId] = useState<string>("");
  const [extractedHtml, setExtractedHtml] = useState<string>("");
  const [pageTimings, setPageTimings] = useState<OcrStatus["page_timings"]>([]);

  const [excelSheets, setExcelSheets] = useState<ExcelSheet[] | null>(null);
  const [isExcelPreviewOpen, setIsExcelPreviewOpen] = useState(false);
  const [isExcelLoading, setIsExcelLoading] = useState(false);

  const [dbName, setDbName] = useState("OCR_DB");
  const [dbTableBase, setDbTableBase] = useState("INVO_FILM");
  const [dbMessage, setDbMessage] = useState<string>("");
  const [dbMessageError, setDbMessageError] = useState<boolean>(false);
  const [isUploadingDb, setIsUploadingDb] = useState(false);

  const excelPreviewRef = useRef<HTMLDivElement | null>(null);

  // realtime elapsed time while OCR is running
  const [elapsedSeconds, setElapsedSeconds] = useState<number | null>(null);
  const ocrStartRef = useRef<number | null>(null);
  const elapsedTimerRef = useRef<number | null>(null);

  // image lightbox
  const [lightboxOpen, setLightboxOpen] = useState(false);
  const [lightboxScale, setLightboxScale] = useState(1);
  const [lightboxBaseScale, setLightboxBaseScale] = useState(1);
  const lightboxImgRef = useRef<HTMLImageElement | null>(null);
  const lightboxScrollRef = useRef<HTMLDivElement | null>(null);
  const lightboxWrapRef = useRef<HTMLDivElement | null>(null);
  const ocrResultSectionRef = useRef<HTMLElement | null>(null);

  // -------- File handling --------

  function handleFileChange(file: File | null) {
    if (!file) {
      setSelectedFile(null);
      setSelectedFileName("ยังไม่ได้เลือกไฟล์");
      setImagePreviewUrl(null);
      return;
    }
    setSelectedFile(file);
    setSelectedFileName(file.name);
    const isImage = file.type.startsWith("image/") || /\.(png|jpe?g)$/i.test(file.name);
    if (isImage) {
      const url = URL.createObjectURL(file);
      setImagePreviewUrl(url);
    } else {
      setImagePreviewUrl(null);
    }
  }

  // -------- OCR start + polling --------

  async function startOcr() {
    if (!selectedFile) {
      alert("กรุณาเลือกไฟล์ก่อน");
      return;
    }
    setIsStarting(true);
    // เริ่มจับเวลาแบบ realtime
    ocrStartRef.current = Date.now();
    setElapsedSeconds(0);
    if (elapsedTimerRef.current !== null) {
      window.clearInterval(elapsedTimerRef.current);
    }
    elapsedTimerRef.current = window.setInterval(() => {
      if (ocrStartRef.current != null) {
        setElapsedSeconds(Math.floor((Date.now() - ocrStartRef.current) / 1000));
      }
    }, 1000);
    setStatus(null);
    setResultId("");
    try {
      const formData = new FormData();
      formData.append("pdf_file", selectedFile);
      const resp = await fetch(`${API_BASE}/ocr/start`, { method: "POST", body: formData });
      const data = await resp.json();
      if (!resp.ok || !data.ok) {
        throw new Error(data.error || "ไม่สามารถเริ่ม OCR ได้");
      }
      setJobId(data.job_id);
    } catch (err: any) {
      alert(err.message || "ไม่สามารถเริ่ม OCR ได้");
    } finally {
      setIsStarting(false);
    }
  }

  useEffect(() => {
    if (!jobId) return;
    let cancelled = false;

    async function poll() {
      try {
        while (!cancelled) {
          const resp = await fetch(`${API_BASE}/ocr/status/${jobId}`);
          const data: OcrStatus = await resp.json();
          if (!resp.ok || !data.ok) {
            alert(data.error || "อ่านสถานะ OCR ไม่ได้");
            break;
          }
          if (cancelled) break;
          setStatus(data);
          setPageTimings(data.page_timings || []);
          if (data.status === "completed" || data.status === "error") {
            if (data.status === "completed" && data.result && data.result_id) {
              setResultId(data.result_id);
              setExtractedHtml(data.result.extracted_html || "");
            }
            break;
          }
          await new Promise((r) => setTimeout(r, 1000));
        }
      } finally {
        if (ocrStartRef.current != null) {
          setElapsedSeconds(Math.floor((Date.now() - ocrStartRef.current) / 1000));
        }
        if (elapsedTimerRef.current !== null) {
          window.clearInterval(elapsedTimerRef.current);
          elapsedTimerRef.current = null;
        }
      }
    }

    poll();
    return () => {
      cancelled = true;
    };
  }, [jobId]);

  // เลื่อนไปที่ผลลัพธ์ OCR เมื่อ OCR เสร็จ (เหมือนใน HTML เดิม)
  useEffect(() => {
    if (!resultId) return;
    const el = ocrResultSectionRef.current;
    if (!el) return;
    const t = requestAnimationFrame(() => {
      el.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    return () => cancelAnimationFrame(t);
  }, [resultId]);

  // -------- Excel preview & editing --------

  async function openExcelPreview() {
    if (!resultId) {
      setIsExcelPreviewOpen(true);
      setExcelSheets(null);
      return;
    }
    setIsExcelPreviewOpen(true);
    setIsExcelLoading(true);
    try {
      const formData = new FormData();
      formData.append("result_id", resultId);
      const resp = await fetch(`${API_BASE}/preview/excel`, {
        method: "POST",
        body: formData,
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || "โหลด preview ไม่สำเร็จ");
      setExcelSheets((data.sheets || []) as ExcelSheet[]);
    } catch (err: any) {
      alert(err.message || "โหลด preview ไม่สำเร็จ");
    } finally {
      setIsExcelLoading(false);
    }
  }

  function getEditedSheetsFromPreview(): { sheets: ExcelSheet[] } | null {
    const root = excelPreviewRef.current;
    if (!root) return null;
    const sections = root.querySelectorAll<HTMLDivElement>(".excel-preview-sheet");
    if (sections.length === 0) return null;
    const sheets: ExcelSheet[] = [];
    sections.forEach((section) => {
      const table = section.querySelector("table");
      const name = section.dataset.sheetName || "Sheet";
      if (!table) return;
      const rows: string[][] = [];
      const headerRow = table.querySelector("thead tr");
      if (headerRow) {
        const headerCells = headerRow.querySelectorAll<HTMLElement>("th");
        rows.push(Array.from(headerCells).map((el) => (el.textContent || "").trim()));
      }
      const bodyRows = table.querySelectorAll<HTMLTableRowElement>("tbody tr");
      bodyRows.forEach((tr) => {
        const cells = tr.querySelectorAll<HTMLElement>("td");
        rows.push(Array.from(cells).map((el) => (el.textContent || "").trim()));
      });
      sheets.push({ name, rows });
    });
    return sheets.length ? { sheets } : null;
  }

  async function downloadExcel() {
    if (!resultId) {
      alert("ยังไม่มีผลลัพธ์ OCR สำหรับดาวน์โหลด");
      return;
    }
    const formData = new FormData();
    formData.append("result_id", resultId);
    const edited = getEditedSheetsFromPreview();
    if (edited) formData.append("sheets_json", JSON.stringify(edited));
    try {
      const resp = await fetch(`${API_BASE}/download/excel`, {
        method: "POST",
        body: formData,
      });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err.error || "ดาวน์โหลดไม่สำเร็จ");
      }
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "ocr-tables.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err: any) {
      alert(err.message || "ดาวน์โหลดไม่สำเร็จ");
    }
  }

  // -------- Upload to DB --------

  function setDbUploadMessage(message: string, isError = false) {
    setDbMessage(message);
    setDbMessageError(isError);
  }

  async function uploadToDb() {
    if (!resultId) {
      setDbUploadMessage("ยังไม่มีผลลัพธ์ OCR สำหรับอัพโหลด", true);
      return;
    }
    if (!dbTableBase) {
      setDbUploadMessage("กรุณากรอกชื่อตารางฐานข้อมูล", true);
      return;
    }
    setIsUploadingDb(true);
    setDbUploadMessage("กำลังตรวจสอบเลขที่ในฐานข้อมูล...");
    const body = new URLSearchParams();
    body.append("result_id", resultId);
    body.append("table_name", dbTableBase);
    body.append("db_name", dbName || "ExcelTtbDB");

    // duplicate check
    let doUpload = true;
    try {
      const checkResp = await fetch(`${API_BASE}/upload/db/check`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      });
      const checkData = await checkResp.json();
      if (!checkResp.ok || !checkData.ok) {
        setDbUploadMessage(checkData.error || "ตรวจสอบไม่สำเร็จ", true);
        doUpload = false;
      } else if (checkData.exists && checkData.doc_no) {
        const msg = `พบเลขที่ "${checkData.doc_no}" ในฐานข้อมูลแล้ว ต้องการเขียนทับข้อมูลเดิมหรือไม่?`;
        if (!window.confirm(msg)) {
          setDbUploadMessage("ยกเลิกการอัพโหลด");
          doUpload = false;
        }
      }
    } catch (err: any) {
      setDbUploadMessage(err.message || "ตรวจสอบไม่สำเร็จ", true);
      doUpload = false;
    }

    if (!doUpload) {
      setIsUploadingDb(false);
      return;
    }

    setDbUploadMessage("กำลังอัพโหลดข้อมูลลงฐานข้อมูล...");
    const uploadBody = new URLSearchParams(body.toString());
    if (isExcelPreviewOpen) {
      const edited = getEditedSheetsFromPreview();
      if (edited) uploadBody.append("edited_sheets_json", JSON.stringify(edited));
    }

    try {
      const resp = await fetch(`${API_BASE}/upload/db`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: uploadBody.toString(),
      });
      const data = await resp.json();
      if (!resp.ok || !data.ok) {
        throw new Error(data.error || "อัพโหลดฐานข้อมูลไม่สำเร็จ");
      }
      const info = data.upload_info || {};
      const parts = [
        `อัพโหลดสำเร็จ: ${info.db_name || ""} → ${info.header_table || ""}, ${info.detail_table || ""}, ${info.total_table || ""}`,
      ];
      if (
        info.header_rows != null &&
        info.detail_rows != null &&
        info.total_rows != null
      ) {
        parts.push(
          `(header ${info.header_rows} แถว, detail ${info.detail_rows} แถว, total ${info.total_rows} แถว)`,
        );
      } else {
        parts.push(`(${info.rows || 0} แถว)`);
      }
      setDbUploadMessage(parts.join(" "));
    } catch (err: any) {
      setDbUploadMessage(err.message || "อัพโหลดฐานข้อมูลไม่สำเร็จ", true);
    } finally {
      setIsUploadingDb(false);
    }
  }

  // -------- Lightbox helpers --------

  function getLightboxFitScale(): number {
    const scroll = lightboxScrollRef.current;
    const img = lightboxImgRef.current;
    if (!scroll || !img) return 1;
    const nw = img.naturalWidth;
    const nh = img.naturalHeight;
    if (!nw || !nh) return 1;
    const w = scroll.clientWidth - 32;
    const h = scroll.clientHeight - 32;
    if (w <= 0 || h <= 0) return 1;
    const scaleByW = w / nw;
    const scaleByH = h / nh;
    return Math.min(1, scaleByW, scaleByH);
  }

  function applyLightboxZoom(scale: number) {
    const scroll = lightboxScrollRef.current;
    const img = lightboxImgRef.current;
    const wrap = lightboxWrapRef.current;
    const MIN_ZOOM = 0.25;
    const MAX_ZOOM = 4;
    const newScale = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, scale));
    setLightboxScale(newScale);
    if (!img || !wrap || !scroll) return;
    const nw = img.naturalWidth;
    const nh = img.naturalHeight;
    if (nw && nh) {
      wrap.style.width = `${nw * newScale}px`;
      wrap.style.height = `${nh * newScale}px`;
    }
    img.style.transform = `scale(${newScale})`;
    img.style.width = `${img.naturalWidth || 0}px`;
    img.style.height = `${img.naturalHeight || 0}px`;

    requestAnimationFrame(() => {
      const wrapW = wrap.offsetWidth;
      const wrapH = wrap.offsetHeight;
      const overflowX = wrapW > scroll.clientWidth;
      const overflowY = wrapH > scroll.clientHeight;
      const maxScrollLeft = Math.max(0, scroll.scrollWidth - scroll.clientWidth);
      if (overflowX) {
        scroll.classList.add("content-overflow");
      } else {
        scroll.classList.remove("content-overflow");
      }
      scroll.scrollLeft = overflowX ? maxScrollLeft / 2 : 0;
      scroll.scrollTop = overflowY ? 0 : 0;
    });
  }

  function openLightbox() {
    if (!imagePreviewUrl) return;
    setLightboxOpen(true);
    setTimeout(() => {
      const img = lightboxImgRef.current;
      if (!img) return;
      img.onload = () => {
        const fit = getLightboxFitScale();
        setLightboxBaseScale(fit);
        applyLightboxZoom(fit);
      };
      img.src = imagePreviewUrl;
      if (img.complete && img.naturalWidth) {
        const fit = getLightboxFitScale();
        setLightboxBaseScale(fit);
        applyLightboxZoom(fit);
      }
    }, 0);
    document.body.style.overflow = "hidden";
  }

  function closeLightbox() {
    setLightboxOpen(false);
    document.body.style.overflow = "";
  }

  // drag-to-pan in lightbox
  useEffect(() => {
    const scrollEl = lightboxScrollRef.current;
    if (!scrollEl) return;
    const el = scrollEl;
    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let startScrollLeft = 0;
    let startScrollTop = 0;

    function onMouseDown(e: MouseEvent) {
      if (e.button !== 0) return;
      if (!el.classList.contains("content-overflow")) return;
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      startScrollLeft = el.scrollLeft;
      startScrollTop = el.scrollTop;
      el.classList.add("cursor-grabbing");
      e.preventDefault();
    }

    function onMouseMove(e: MouseEvent) {
      if (!isDragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      el.scrollLeft = startScrollLeft - dx;
      el.scrollTop = startScrollTop - dy;
    }

    function onMouseUp() {
      if (!isDragging) return;
      isDragging = false;
      el.classList.remove("cursor-grabbing");
    }

    el.addEventListener("mousedown", onMouseDown);
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    return () => {
      el.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
    };
  }, [lightboxOpen]);

  // -------- Render helpers --------

  const currentStep = status?.current_step || 0;
  const totalSteps = status?.total_steps || 0;
  const progressPercent =
    totalSteps > 0 ? Math.round((currentStep / totalSteps) * 100) : status ? 100 : 0;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <div className="mx-auto max-w-6xl p-4 md:p-8">
        <div className="mb-6 rounded-2xl bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white shadow-lg">
          <h1 className="text-2xl font-bold md:text-3xl">
            OCR ใบส่งของ/ใบกำกับ — Typhoon
          </h1>
          <p className="mt-2 text-sm text-blue-100 md:text-base">
            ลากไฟล์ PDF หรือรูปสแกน (PNG/JPG) มาวางได้เลย — รองรับฟอร์แมตใบส่งของ บริษัท
            เอส พี บ้านคาร์แคร์ จำกัด
          </p>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          {/* Upload panel */}
          <UploadPanel
            selectedFile={selectedFile}
            selectedFileName={selectedFileName}
            handleFileChange={handleFileChange}
            startOcr={startOcr}
            isStarting={isStarting}
            status={status}
            currentStep={currentStep}
            totalSteps={totalSteps}
            progressPercent={progressPercent}
            elapsedSeconds={elapsedSeconds}
            pageTimings={pageTimings}
          />

          {/* Image preview panel */}
          <ImagePreviewPanel
            imagePreviewUrl={imagePreviewUrl}
            selectedFileName={selectedFileName}
            openLightbox={openLightbox}
          />
        </div>

        {/* OCR result + Excel/DB controls */}
        <OcrResultPanel
          extractedHtml={extractedHtml}
          ocrResultSectionRef={ocrResultSectionRef}
          resultId={resultId}
          openExcelPreview={openExcelPreview}
          isExcelPreviewOpen={isExcelPreviewOpen}
          setIsExcelPreviewOpen={setIsExcelPreviewOpen}
          downloadExcel={downloadExcel}
          dbName={dbName}
          setDbName={setDbName}
          dbTableBase={dbTableBase}
          setDbTableBase={setDbTableBase}
          uploadToDb={uploadToDb}
          isUploadingDb={isUploadingDb}
          dbMessage={dbMessage}
          dbMessageError={dbMessageError}
          isExcelLoading={isExcelLoading}
          excelPreviewRef={excelPreviewRef}
          excelSheets={excelSheets}
        />
      </div>

      {/* Lightbox */}
      <Lightbox
        lightboxOpen={lightboxOpen}
        closeLightbox={closeLightbox}
        applyLightboxZoom={applyLightboxZoom}
        lightboxScale={lightboxScale}
        lightboxBaseScale={lightboxBaseScale}
        lightboxScrollRef={lightboxScrollRef}
        lightboxWrapRef={lightboxWrapRef}
        lightboxImgRef={lightboxImgRef}
      />
    </div>
  );
};

export default App;
