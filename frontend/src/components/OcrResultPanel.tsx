import React, { useState, useEffect } from "react";
import type { ExcelSheet } from "../types";

interface Document {
    page_number: number;
    วันที่: string;
    เลขที่: string;
    พนักงานขาย: string;
    กำหนดชำระเงิน: string;
    ครบกำหนดวันที่: string;
    รวมเงิน: string;
    ภาษีมูลค่าเพิ่ม: string;
    รวมสุทธิ: string;
    has_error: boolean;
    error: string;
}

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
    pageTimings: any[];
    elapsedSeconds: number | null;
    // PDF Preview props
    pdfPreviewUrl?: string | null;
    isPdfFile?: boolean;
    // OCR status for real-time updates
    ocrStatus?: string | null;
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
    pageTimings,
    elapsedSeconds,
    // PDF Preview props
    pdfPreviewUrl,
    isPdfFile,
    // OCR status for real-time updates
    ocrStatus,
}) => {
    const [documents, setDocuments] = useState<Document[]>([]);
    const [isLoadingDocuments, setIsLoadingDocuments] = useState(false);
    const [selectedPage, setSelectedPage] = useState<number | null>(null);
    const [pageDetails, setPageDetails] = useState<any>(null);
    const [showModal, setShowModal] = useState(false);
    const [viewMode, setViewMode] = useState<'list' | 'full'>('list'); // เริ่มต้นที่ 'list' เสมอ

    // สถานะสำหรับแก้ไขข้อมูลใน modal
    const [editedHeader, setEditedHeader] = useState<Record<string, string>>({});
    const [editedDetail, setEditedDetail] = useState<any[]>([]);
    const [editedTotal, setEditedTotal] = useState<Record<string, string>>({});
    const [hasEdits, setHasEdits] = useState(false);

    // ดึงข้อมูลรายการเอกสารเมื่อมี resultId
    useEffect(() => {
        if (resultId) {
            loadDocuments();
        }
    }, [resultId]);

    // รีโหลดข้อมูลแบบ real-time ขณะ OCR กำลังทำงาน
    useEffect(() => {
        console.log('[DEBUG] Real-time useEffect triggered - resultId:', resultId, 'ocrStatus:', ocrStatus);
        if (!resultId || ocrStatus !== 'running') {
            console.log('[DEBUG] Skipping real-time refresh - no resultId or status not running');
            return;
        }
        
        console.log('[DEBUG] Starting real-time auto-refresh interval');
        // รีโหลดทุก 5 วินาที ขณะ OCR ยังทำงานอยู่ (เพื่อลดการกระพริบสำหรับไฟล์หลายหน้า)
        const interval = setInterval(() => {
            console.log('[DEBUG] Auto-refreshing documents...');
            loadDocuments();
        }, 5000);
        
        return () => {
            console.log('[DEBUG] Clearing real-time refresh interval');
            clearInterval(interval);
        };
    }, [resultId, ocrStatus]);

    // รีโหลดข้อมูลครั้งสุดท้ายเมื่อ OCR เสร็จสมบูรณ์ (เพื่อรับข้อมูล merged ที่ถูกต้อง)
    useEffect(() => {
        if (resultId && ocrStatus === 'completed') {
            console.log('[DEBUG] OCR completed - loading final documents with merged data');
            loadDocuments();
        }
    }, [resultId, ocrStatus]);

    async function loadDocuments() {
        if (!resultId) return;
        
        console.log('Loading documents for resultId:', resultId);
        setIsLoadingDocuments(true);
        try {
            const resp = await fetch(`/api/list-documents?result_id=${resultId}`);
            console.log('Response status:', resp.status);
            const data = await resp.json();
            console.log('Response data:', data);
            
            if (!resp.ok || !data.ok) {
                console.error('ไม่สามารถโหลดข้อมูลได้:', data.error);
                return;
            }
            
            console.log('Documents loaded:', data.documents);
            setDocuments(data.documents || []);
        } catch (error) {
            console.error('เกิดข้อผิดพลาด:', error);
        } finally {
            setIsLoadingDocuments(false);
        }
    }

    async function viewPageDetails(pageNumber: number) {
        if (!resultId) return;
        
        console.log(`[DEBUG] Opening page details for page ${pageNumber}`);
        
        try {
            const resp = await fetch(`/api/page-details?result_id=${resultId}&page_number=${pageNumber}`);
            const data = await resp.json();
            
            console.log(`[DEBUG] Received page details:`, data);
            
            if (!resp.ok || !data.ok) {
                alert(`ไม่สามารถโหลดข้อมูลหน้า ${pageNumber} ได้: ${data.error}`);
                return;
            }
            
            console.log(`[DEBUG] Setting selectedPage to ${pageNumber}, pageDetails page_number=${data.page_result?.page_number}`);
            
            setPageDetails(data.page_result);
            setSelectedPage(pageNumber);
            // เริ่มต้นข้อมูลที่แก้ไขจากข้อมูลเดิม
            setEditedHeader(data.page_result?.header || {});
            // ให้แก้ไขได้ทุกแถว/ทุกช่อง (รวมแถวแรกด้วย)
            setEditedDetail(data.page_result?.detail || []);
            setEditedTotal(data.page_result?.total || {});
            setHasEdits(false);
            setShowModal(true);
        } catch (error) {
            alert(`เกิดข้อผิดพลาด: ${error}`);
        }
    }

    function closeModal() {
        setShowModal(false);
        setSelectedPage(null);
        setPageDetails(null);
        // เคลียร์ข้อมูลที่แก้ไข
        setEditedHeader({});
        setEditedDetail([]);
        setEditedTotal({});
        setHasEdits(false);
    }

    // ฟังก์ชันจัดการการแก้ไขข้อมูล
    function handleHeaderChange(key: string, value: string) {
        setEditedHeader(prev => ({ ...prev, [key]: value }));
        setHasEdits(true);
    }

    function handleDetailChange(rowIndex: number, cellIndex: number, value: string) {
        setEditedDetail(prev => {
            const newDetail = [...prev];
            newDetail[rowIndex] = [...newDetail[rowIndex]];
            newDetail[rowIndex][cellIndex] = value;
            return newDetail;
        });
        setHasEdits(true);
    }

    function handleTotalChange(key: string, value: string) {
        setEditedTotal(prev => ({ ...prev, [key]: value }));
        setHasEdits(true);
    }

    function saveEdits() {
        if (!pageDetails) return;
        
        // อัปเดตข้อมูลใน pageDetails
        const updatedPageDetails = {
            ...pageDetails,
            header: editedHeader,
            detail: editedDetail,
            total: editedTotal
        };
        
        // ส่งข้อมูลที่แก้ไขไปยัง API เพื่ออัปเดต
        fetch('/api/update-page-details', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                result_id: resultId,
                page_number: selectedPage,
                page_result: updatedPageDetails
            })
        }).then(resp => resp.json())
          .then(data => {
              if (data.ok) {
                  setPageDetails(updatedPageDetails);
                  setHasEdits(false);
                  alert('บันทึกการแก้ไขสำเร็จ');
                  // รีโหลดรายการเอกสารเพื่ออัปเดต UI
                  loadDocuments();
              } else {
                  alert(`ไม่สามารถบันทึกได้: ${data.error}`);
              }
          })
          .catch(error => {
              alert(`เกิดข้อผิดพลาด: ${error}`);
          });
    }

    if (!extractedHtml && !resultId) return null;

    console.log('OcrResultPanel rendered - resultId:', resultId, 'extractedHtml:', extractedHtml);

    return (
        <section
            ref={ocrResultSectionRef}
            className="mt-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"
        >
            <div className="flex flex-wrap items-center justify-between gap-2">
                <h2 className="text-xl font-semibold text-slate-900">
                    สรุปเอกสาร
                </h2>
                
                {/* ส่วนอัปโหลดลงฐานข้อมูล */}
                {resultId && (
                    <div className="flex items-center gap-2">
                        <input
                            type="text"
                            placeholder="ชื่อฐานข้อมูล"
                            value={dbName}
                            onChange={(e) => setDbName(e.target.value)}
                            className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                        />
                        <input
                            type="text"
                            placeholder="ชื่อตาราง"
                            value={dbTableBase}
                            onChange={(e) => setDbTableBase(e.target.value)}
                            className="rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                        />
                        <button
                            onClick={uploadToDb}
                            disabled={isUploadingDb}
                            className="rounded-lg bg-green-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-green-700 disabled:cursor-not-allowed disabled:bg-green-400"
                        >
                            {isUploadingDb ? 'กำลังอัปโหลด...' : 'อัปโหลดลงฐานข้อมูล'}
                        </button>
                        {dbMessage && (
                            <div className={`text-sm ${dbMessageError ? 'text-red-600' : 'text-green-600'}`}>
                                {dbMessage}
                            </div>
                        )}
                    </div>
                )}
            </div>

            {/* แสดงเวลาที่ใช้ */}
            {elapsedSeconds !== null && (
                <div className="mt-2 text-sm text-slate-600">
                    ใช้เวลาทั้งหมด {elapsedSeconds} วินาที
                </div>
            )}

            {/* มุมมองรายการ */}
            {resultId && (
                <div className="mt-4">
                    {isLoadingDocuments ? (
                        <div className="text-center py-4">
                            <div className="text-slate-600">กำลังโหลดข้อมูล...</div>
                        </div>
                    ) : documents.length > 0 ? (
                        <div className="overflow-x-auto">
                            <table className="min-w-full border-2 border-slate-300 bg-white shadow-lg">
                                <thead className="bg-gradient-to-r from-blue-50 to-indigo-50 border-b-2 border-slate-300">
                                    <tr>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200 min-w-[80px]">หน้า</th>
                                        <th className="px-8 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200 min-w-[140px]">วันที่</th>
                                        <th className="px-8 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200 min-w-[180px]">เลขที่</th>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200">พนักงานขาย</th>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200">กำหนดชำระเงิน</th>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200">ครบกำหนดวันที่</th>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200">รวมเงิน</th>
                                        <th className="px-4 py-3 text-left text-xs font-bold text-slate-800 uppercase tracking-wider border-r border-slate-200">รวมสุทธิ</th>
                                        <th className="px-4 py-3 text-center text-xs font-bold text-slate-800 uppercase tracking-wider">ดำเนินการ</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-200">
                                    {documents.map((doc, index) => (
                                        <tr key={doc.page_number} className={`${doc.has_error ? 'bg-red-50' : index % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-blue-50 transition-colors`}>
                                            <td className="px-4 py-3 text-sm font-semibold text-slate-900 border-r border-slate-200">{doc.page_number}</td>
                                            <td className="px-8 py-3 text-sm text-slate-900 border-r border-slate-200 whitespace-nowrap">{doc.วันที่ || '-'}</td>
                                            <td className="px-8 py-3 text-sm font-bold text-slate-900 border-r border-slate-200 whitespace-nowrap">
                                                {doc.เลขที่ || '-'}
                                                {doc.has_error && (
                                                    <span className="ml-2 px-2 py-1 text-xs bg-red-100 text-red-800 rounded-full">มีข้อผิดพลาด</span>
                                                )}
                                            </td>
                                            <td className="px-4 py-3 text-sm text-slate-900 border-r border-slate-200">{doc.พนักงานขาย || '-'}</td>
                                            <td className="px-4 py-3 text-sm text-slate-900 border-r border-slate-200">{doc.กำหนดชำระเงิน || '-'}</td>
                                            <td className="px-4 py-3 text-sm text-slate-900 border-r border-slate-200">{doc.ครบกำหนดวันที่ || '-'}</td>
                                            <td className="px-4 py-3 text-sm font-semibold text-slate-900 border-r border-slate-200">{doc.รวมเงิน || '-'}</td>
                                            <td className="px-4 py-3 text-sm font-bold text-emerald-700 border-r border-slate-200">{doc.รวมสุทธิ || '-'}</td>
                                            <td className="px-4 py-3 text-sm text-center">
                                                <button
                                                    onClick={() => viewPageDetails(doc.page_number)}
                                                    className="px-4 py-2 text-xs bg-gradient-to-r from-blue-600 to-indigo-600 text-white rounded-lg hover:from-blue-700 hover:to-indigo-700 transition-all transform hover:scale-105 shadow-md"
                                                >
                                                    ดูรายละเอียด
                                                </button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    ) : (
                        <div className="text-center py-4">
                            <div className="text-slate-500">ไม่พบข้อมูลเอกสาร</div>
                        </div>
                    )}
                </div>
            )}

            {/* Modal สำหรับดูรายละเอียดหน้า */}
            {showModal && pageDetails && (
                <div className="fixed inset-0 bg-black bg-opacity-70 flex items-center justify-center z-50 p-2">
                    <div className="bg-white rounded-xl w-[98vw] h-[96vh] overflow-hidden shadow-2xl flex flex-col">
                        <div className="sticky top-0 bg-gradient-to-r from-blue-600 to-indigo-600 text-white border-b border-slate-200 p-4 flex justify-between items-center rounded-t-xl">
                            <h3 className="text-lg font-bold">รายละเอียดหน้า {selectedPage}</h3>
                            <div className="flex items-center gap-2">
                                {hasEdits && (
                                    <button
                                        onClick={saveEdits}
                                        className="px-4 py-2 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-sm font-semibold transition-colors flex items-center gap-2"
                                    >
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7"></path>
                                        </svg>
                                        บันทึกการแก้ไข
                                    </button>
                                )}
                                <button
                                    onClick={closeModal}
                                    className="text-white hover:bg-white hover:bg-opacity-20 rounded-lg p-2 transition-colors"
                                >
                                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"></path>
                                    </svg>
                                </button>
                            </div>
                        </div>
                        <div className="p-0 flex-1 overflow-hidden">
                            {/* แสดง PDF + ข้อมูลแก้ไขแบบ side-by-side */}
                            <div className="flex flex-col lg:flex-row h-full">
                                {/* ด้านซ้าย: PDF Preview - 2/5 ส่วน */}
                                {isPdfFile && pdfPreviewUrl && (
                                    <div className="lg:w-2/5 w-full border-r border-slate-200 bg-slate-100 p-4 overflow-hidden flex flex-col">
                                        <h4 className="text-sm font-bold text-slate-700 mb-3 flex items-center gap-2">
                                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path>
                                            </svg>
                                            หน้า PDF ที่ {selectedPage}
                                        </h4>
                                        <div className="bg-white rounded-lg border border-slate-200 shadow-sm flex-1 overflow-hidden">
                                            <iframe
                                                src={`${pdfPreviewUrl}#page=${selectedPage}&zoom=page-fit&toolbar=1&navpanes=0&scrollbar=1&pagemode=none`}
                                                width="100%"
                                                height="100%"
                                                style={{ border: 'none' }}
                                                title={`PDF Page ${selectedPage}`}
                                            />
                                        </div>
                                    </div>
                                )}
                                
                                {/* ด้านขวา: แบบฟอร์มแก้ไข - 3/5 ส่วน */}
                                <div className={`${isPdfFile && pdfPreviewUrl ? 'lg:w-3/5' : 'w-full'} w-full p-6 overflow-y-auto bg-white`}>
                                    <div className="space-y-6">
                                        {/* ส่วนหัวเอกสาร - แบบแก้ไขได้ */}
                                        {editedHeader && Object.keys(editedHeader).length > 0 && (
                                            <div className="bg-gradient-to-r from-slate-50 to-blue-50 rounded-lg p-6 border border-slate-200">
                                                <h4 className="text-lg font-bold text-slate-900 mb-4 pb-2 border-b-2 border-blue-200">ข้อมูลเอกสาร (แก้ไขได้)</h4>
                                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                                    {Object.entries(editedHeader).map(([key, value]: [string, any]) => (
                                                        <div key={key} className="bg-white rounded-lg p-3 border border-slate-200 shadow-sm">
                                                            <label className="text-xs font-semibold text-slate-600 uppercase tracking-wider mb-1 block">{key}</label>
                                                            <input
                                                                type="text"
                                                                value={String(value || '')}
                                                                onChange={(e) => handleHeaderChange(key, e.target.value)}
                                                                className="w-full text-sm font-bold text-slate-900 border border-slate-300 rounded px-2 py-1 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                                                            />
                                                        </div>
                                                    ))}
                                                </div>
                                            </div>
                                        )}

                                        {/* ตารางรายการสินค้า - แบบแก้ไขได้ */}
                                        {pageDetails.detail && pageDetails.detail.length > 0 && (
                                            <div className="bg-white rounded-lg p-6 border-2 border-slate-200 shadow-lg">
                                                <h4 className="text-lg font-bold text-slate-900 mb-4 pb-2 border-b-2 border-slate-200">รายการสินค้า (แก้ไขได้)</h4>
                                                <div className="overflow-x-auto">
                                                    <table className="min-w-full border-2 border-slate-300">
                                                        <tbody className="divide-y divide-slate-200">
                                                            {editedDetail.map((row: any[], rowIndex: number) => (
                                                                <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-white' : 'bg-slate-50'}>
                                                                    {row.map((cell: any, cellIndex: number) => (
                                                                        <td key={cellIndex} className="px-2 py-2 text-sm text-slate-900 border-r border-slate-200">
                                                                            <input
                                                                                type="text"
                                                                                value={String(cell || '')}
                                                                                onChange={(e) => handleDetailChange(rowIndex, cellIndex, e.target.value)}
                                                                                className="w-full text-sm text-slate-900 border border-slate-300 rounded px-2 py-1 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
                                                                            />
                                                                        </td>
                                                                    ))}
                                                                </tr>
                                                            ))}
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>
                                        )}

                                        {/* สรุปยอดเงิน - แบบแก้ไขได้ */}
                                        {editedTotal && Object.keys(editedTotal).length > 0 && (
                                            <div className="bg-gradient-to-r from-emerald-50 to-green-50 rounded-lg p-6 border border-emerald-200">
                                                <h4 className="text-lg font-bold text-slate-900 mb-4 pb-2 border-b-2 border-emerald-200">สรุปยอดเงิน (แก้ไขได้)</h4>
                                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                                    {Object.entries(editedTotal).map(([key, value]: [string, any]) => (
                                                        <div key={key} className="bg-white rounded-lg p-4 border border-emerald-200 shadow-sm">
                                                            <label className="text-xs font-semibold text-emerald-600 uppercase tracking-wider mb-1 block">{key}</label>
                                                            <input
                                                                type="text"
                                                                value={String(value || '')}
                                                                onChange={(e) => handleTotalChange(key, e.target.value)}
                                                                className="w-full text-lg font-bold text-emerald-700 border border-emerald-300 rounded px-2 py-1 focus:border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                                                            />
                                                        </div>
                                                    ))}
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            )}

            </section>
    );
};
