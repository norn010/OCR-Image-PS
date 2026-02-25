import React from "react";

export interface ImagePreviewPanelProps {
    imagePreviewUrl: string | null;
    selectedFileName: string;
    openLightbox: () => void;
}

export const ImagePreviewPanel: React.FC<ImagePreviewPanelProps> = ({
    imagePreviewUrl,
    selectedFileName,
    openLightbox,
}) => {
    return (
        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm min-h-[20rem]">
            <h2 className="mb-4 text-lg font-semibold text-slate-900">รูปสแกน (Preview)</h2>
            {!imagePreviewUrl && (
                <div className="flex min-h-[10rem] items-center justify-center rounded-xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                    ยังไม่มีรูปตัวอย่าง
                </div>
            )}
            {imagePreviewUrl && (
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
                        {selectedFileName}
                    </div>
                    <button
                        type="button"
                        onClick={openLightbox}
                        className="block w-full cursor-zoom-in rounded border border-slate-300 bg-white focus:outline-none focus:ring-2 focus:ring-blue-400"
                    >
                        {/* eslint-disable-next-line jsx-a11y/alt-text */}
                        <img
                            src={imagePreviewUrl}
                            className="mx-auto max-h-80 w-full max-w-md rounded object-contain"
                        />
                    </button>
                    <p className="mt-1 text-center text-xs text-slate-500">คลิกเพื่อดูรูปใหญ่</p>
                </div>
            )}
        </section>
    );
};
