import React from "react";

export interface LightboxProps {
    lightboxOpen: boolean;
    closeLightbox: () => void;
    applyLightboxZoom: (scale: number) => void;
    lightboxScale: number;
    lightboxBaseScale: number;
    lightboxScrollRef: React.RefObject<HTMLDivElement | null>;
    lightboxWrapRef: React.RefObject<HTMLDivElement | null>;
    lightboxImgRef: React.RefObject<HTMLImageElement | null>;
}

export const Lightbox: React.FC<LightboxProps> = ({
    lightboxOpen,
    closeLightbox,
    applyLightboxZoom,
    lightboxScale,
    lightboxBaseScale,
    lightboxScrollRef,
    lightboxWrapRef,
    lightboxImgRef,
}) => {
    if (!lightboxOpen) return null;

    return (
        <div
            className="fixed inset-0 z-50 flex flex-col bg-black/90"
            onClick={(e) => {
                if (e.target === e.currentTarget) closeLightbox();
            }}
        >
            <div className="flex shrink-0 items-center justify-between gap-2 border-b border-white/20 bg-black/50 px-4 py-2">
                <div className="flex items-center gap-2">
                    <button
                        type="button"
                        className="rounded-lg bg-white/90 px-3 py-1.5 text-sm font-bold text-slate-800 shadow hover:bg-white"
                        onClick={() => applyLightboxZoom(lightboxScale / 1.25)}
                    >
                        −
                    </button>
                    <span className="min-w-[4rem] text-center text-sm font-medium text-white">
                        {Math.round((lightboxScale / (lightboxBaseScale || 1)) * 100)}%
                    </span>
                    <button
                        type="button"
                        className="rounded-lg bg-white/90 px-3 py-1.5 text-sm font-bold text-slate-800 shadow hover:bg-white"
                        onClick={() => applyLightboxZoom(lightboxScale * 1.25)}
                    >
                        +
                    </button>
                    <button
                        type="button"
                        className="rounded-lg bg-white/70 px-2 py-1.5 text-xs text-slate-700 hover:bg-white"
                        onClick={() => applyLightboxZoom(1)}
                    >
                        100%
                    </button>
                </div>
                <button
                    type="button"
                    className="rounded-full bg-white/90 p-2 text-slate-700 shadow hover:bg-white"
                    aria-label="ปิด"
                    onClick={closeLightbox}
                >
                    ×
                </button>
            </div>
            <div
                ref={lightboxScrollRef}
                className="image-lightbox-scroll flex min-h-0 flex-1 items-center justify-center overflow-auto p-4"
            >
                <div ref={lightboxWrapRef} className="inline-block flex-shrink-0">
                    {/* eslint-disable-next-line jsx-a11y/alt-text */}
                    <img
                        ref={lightboxImgRef}
                        className="block origin-top-left object-contain"
                        draggable={false}
                    />
                </div>
            </div>
        </div>
    );
};
