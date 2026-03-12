import React from "react";

export interface PdfPreviewPanelProps {
    pdfPreviewUrl: string | null;
    selectedFileName: string;
    openLightbox: () => void;
}

export const PdfPreviewPanel: React.FC<PdfPreviewPanelProps> = ({
    pdfPreviewUrl,
    selectedFileName,
}) => {
    return (
        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm min-h-[20rem]">
            <h2 className="mb-4 text-lg font-semibold text-slate-900">PDF Preview</h2>
            {!pdfPreviewUrl && (
                <div className="flex min-h-[10rem] items-center justify-center rounded-xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                    ยังไม่มี PDF ตัวอย่าง
                </div>
            )}
            {pdfPreviewUrl && (
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
                        {selectedFileName}
                    </div>
                    
                    <div className="w-full" style={{ height: "500px" }}>
                        <iframe
                            src={pdfPreviewUrl}
                            width="100%"
                            height="100%"
                            style={{ border: "none", borderRadius: "8px" }}
                            title="PDF Preview"
                        />
                    </div>
                    
                    <p className="mt-2 text-center text-xs text-slate-500">
                        ใช้ scroll หรือปุ่มใน PDF viewer เพื่อนำทาง
                    </p>
                </div>
            )}
        </section>
    );
};
