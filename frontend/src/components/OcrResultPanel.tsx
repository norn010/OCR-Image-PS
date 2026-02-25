import React from "react";
import type { ExcelSheet } from "../types";

export interface OcrResultPanelProps {
    extractedHtml: string;
    ocrResultSectionRef: React.RefObject<HTMLElement | null>;
    resultId: string;
    openExcelPreview: () => void;
    isExcelPreviewOpen: boolean;
    setIsExcelPreviewOpen: (open: boolean) => void;
    downloadExcel: () => void;
    dbName: string;
    setDbName: (name: string) => void;
    dbTableBase: string;
    setDbTableBase: (name: string) => void;
    uploadToDb: () => void;
    isUploadingDb: boolean;
    dbMessage: string;
    dbMessageError: boolean;
    isExcelLoading: boolean;
    excelPreviewRef: React.RefObject<HTMLDivElement | null>;
    excelSheets: ExcelSheet[] | null;
}

export const OcrResultPanel: React.FC<OcrResultPanelProps> = ({
    extractedHtml,
    ocrResultSectionRef,
    resultId,
    openExcelPreview,
    isExcelPreviewOpen,
    setIsExcelPreviewOpen,
    downloadExcel,
    dbName,
    setDbName,
    dbTableBase,
    setDbTableBase,
    uploadToDb,
    isUploadingDb,
    dbMessage,
    dbMessageError,
    isExcelLoading,
    excelPreviewRef,
    excelSheets,
}) => {
    if (!extractedHtml) return null;

    return (
        <section
            ref={ocrResultSectionRef}
            className="mt-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"
        >
            <div className="flex flex-wrap items-center justify-between gap-2">
                <h2 className="text-xl font-semibold text-slate-900">ผลลัพธ์ OCR</h2>
            </div>

            {resultId && (
                <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="flex items-center gap-2">
                            <button
                                type="button"
                                onClick={openExcelPreview}
                                className="rounded-lg border border-emerald-600 bg-white px-3 py-2 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-50"
                            >
                                Preview Excel
                            </button>
                            {isExcelPreviewOpen && (
                                <button
                                    type="button"
                                    onClick={downloadExcel}
                                    className="rounded-lg border border-emerald-600 bg-white px-2 py-1 text-xs font-semibold text-emerald-700 hover:bg-emerald-50"
                                >
                                    ดาวน์โหลด Excel
                                </button>
                            )}
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                            <input
                                type="text"
                                value={dbName}
                                onChange={(e) => setDbName(e.target.value)}
                                className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                                placeholder="ชื่อฐานข้อมูล"
                            />
                            <input
                                type="text"
                                value={dbTableBase}
                                onChange={(e) => setDbTableBase(e.target.value)}
                                className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                                placeholder="ชื่อตารางฐาน (เช่น OCR_TTB_WEB)"
                            />
                            <button
                                type="button"
                                onClick={uploadToDb}
                                disabled={isUploadingDb}
                                className="rounded-lg bg-slate-800 px-3 py-2 text-sm font-semibold text-white transition hover:bg-slate-900 disabled:cursor-not-allowed disabled:opacity-70"
                            >
                                {isUploadingDb ? "กำลังอัพโหลด..." : "อัพลงฐานข้อมูล"}
                            </button>
                        </div>
                    </div>

                    {dbMessage && (
                        <div
                            className={`mt-2 text-sm ${dbMessageError ? "text-red-600" : "text-emerald-700"
                                }`}
                        >
                            {dbMessage}
                        </div>
                    )}

                    {isExcelPreviewOpen && (
                        <div className="mt-4 rounded-xl border border-slate-200 bg-white p-4">
                            <div className="mb-2 flex items-center justify-between gap-2">
                                <span className="text-sm font-semibold text-slate-700">
                                    Preview Excel
                                </span>
                                <button
                                    type="button"
                                    onClick={() => setIsExcelPreviewOpen(false)}
                                    className="rounded p-1 text-slate-500 hover:bg-slate-100 hover:text-slate-700"
                                    aria-label="ปิด"
                                >
                                    ×
                                </button>
                            </div>
                            {isExcelLoading && (
                                <div className="py-8 text-center text-sm text-slate-500">
                                    กำลังโหลด...
                                </div>
                            )}
                            <div
                                ref={excelPreviewRef}
                                className="overflow-x-auto text-sm text-slate-800"
                            >
                                {!isExcelLoading && (!excelSheets || excelSheets.length === 0) && (
                                    <p className="text-slate-500">ยังไม่มีข้อมูลใน Excel</p>
                                )}
                                {excelSheets &&
                                    excelSheets.map((sheet) => (
                                        <div
                                            key={sheet.name}
                                            className="mb-6 excel-preview-sheet"
                                            data-sheet-name={sheet.name}
                                        >
                                            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
                                                {sheet.name}
                                            </div>
                                            <table className="min-w-full border-collapse text-xs">
                                                {sheet.rows.length > 0 && (
                                                    <thead>
                                                        <tr>
                                                            {sheet.rows[0].map((cell, idx) => (
                                                                <th
                                                                    key={idx}
                                                                    contentEditable
                                                                    className="border border-slate-300 bg-slate-100 px-2 py-1 text-left font-semibold"
                                                                >
                                                                    {cell}
                                                                </th>
                                                            ))}
                                                        </tr>
                                                    </thead>
                                                )}
                                                {sheet.rows.length > 1 && (
                                                    <tbody>
                                                        {sheet.rows.slice(1).map((row, rIdx) => (
                                                            <tr key={rIdx}>
                                                                {row.map((cell, cIdx) => (
                                                                    <td
                                                                        key={cIdx}
                                                                        contentEditable
                                                                        className="border border-slate-200 px-2 py-1 align-top"
                                                                    >
                                                                        {cell}
                                                                    </td>
                                                                ))}
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                )}
                                            </table>
                                        </div>
                                    ))}
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* แสดง HTML ที่ได้จาก OCR อยู่ถัดจากส่วน Preview/DB */}
            <div className="mt-4 markdown-result overflow-x-auto rounded-xl border border-slate-200 bg-white p-4 leading-relaxed">
                <div
                    dangerouslySetInnerHTML={{
                        __html: extractedHtml,
                    }}
                />
            </div>
        </section>
    );
};
