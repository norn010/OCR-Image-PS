import React from "react";
import type { OcrStatus } from "../types";

export interface UploadPanelProps {
    selectedFile: File | null;
    selectedFileName: string;
    handleFileChange: (file: File | null) => void;
    startOcr: () => void;
    isStarting: boolean;
    status: OcrStatus | null;
    currentStep: number;
    totalSteps: number;
    progressPercent: number;
    elapsedSeconds: number | null;
    pageTimings: OcrStatus["page_timings"];
}

export const UploadPanel: React.FC<UploadPanelProps> = ({
    selectedFile,
    selectedFileName,
    handleFileChange,
    startOcr,
    isStarting,
    status,
    currentStep,
    totalSteps,
    progressPercent,
    elapsedSeconds,
    pageTimings,
}) => {
    return (
        <section className="flex flex-col rounded-2xl border border-slate-200 bg-white p-5 shadow-sm min-h-[26rem]">
            <h2 className="mb-4 text-lg font-semibold text-slate-900">อัปโหลดไฟล์</h2>

            <label
                htmlFor="pdf_file"
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                    e.preventDefault();
                    const file = e.dataTransfer.files?.[0] || null;
                    handleFileChange(file);
                }}
                className="flex flex-1 cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed border-slate-300 bg-slate-50 p-5 text-center transition hover:border-blue-400 hover:bg-blue-50"
            >
                <div className="text-sm font-medium text-slate-700">ลากไฟล์ PDF หรือรูปสแกนมาวางที่นี่</div>
                <div className="mt-1 text-xs text-slate-500">หรือคลิกเพื่อเลือกไฟล์ (.pdf, .png, .jpg)</div>
                <div className="mt-3 rounded-full bg-slate-200 px-3 py-1 text-xs text-slate-700">
                    {selectedFileName}
                </div>
            </label>
            <input
                id="pdf_file"
                type="file"
                accept=".pdf,.png,.jpg,.jpeg"
                className="hidden"
                onChange={(e) => handleFileChange(e.target.files?.[0] || null)}
            />

            <button
                type="button"
                onClick={startOcr}
                disabled={isStarting || !selectedFile}
                className="mt-4 w-full rounded-xl bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-70 focus:outline-none focus:ring-4 focus:ring-blue-200"
            >
                {isStarting ? "กำลัง OCR..." : "เริ่ม OCR"}
            </button>

            {status && (
                <div className="mt-4 rounded-xl border border-blue-200 bg-blue-50 p-3 text-sm">
                    <div className="font-semibold text-blue-700">
                        {status.status === "running"
                            ? status.message || `กำลัง OCR หน้า ${currentStep || 1}/${totalSteps || 1}`
                            : status.status === "completed"
                                ? "OCR เสร็จแล้ว"
                                : status.status === "error"
                                    ? `เกิดข้อผิดพลาด: ${status.error || status.message || ""}`
                                    : status.message || ""}
                    </div>
                    <div className="mt-2 h-2 w-full rounded-full bg-blue-100">
                        <div
                            className="h-2 rounded-full bg-blue-600 transition-all"
                            style={{ width: `${progressPercent}%` }}
                        />
                    </div>
                    <div className="mt-2 text-xs text-blue-700">
                        เวลาที่ใช้: {elapsedSeconds != null ? `${elapsedSeconds} วินาที` : "—"}
                    </div>
                    <div className="mt-2">
                        <div className="text-xs font-semibold text-blue-700">เวลาแต่ละหน้า</div>
                        <ul className="mt-1 max-h-28 overflow-y-auto text-xs text-blue-800">
                            {pageTimings && pageTimings.length
                                ? pageTimings.map((t) => (
                                    <li key={t.page_number}>
                                        หน้า {t.page_number}: {t.elapsed_seconds} วินาที
                                        {(t as any).retry_count > 0 && (
                                            <span className="ml-1 rounded bg-amber-100 px-1 text-amber-700 font-medium">
                                                🔄 retry ×{(t as any).retry_count}
                                            </span>
                                        )}
                                    </li>
                                ))
                                : "ยังไม่มีข้อมูลเวลา"}
                        </ul>
                    </div>
                </div>
            )}
        </section>
    );
};
