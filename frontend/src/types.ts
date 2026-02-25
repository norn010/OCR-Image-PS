export type OcrStatus = {
    ok: boolean;
    status: "pending" | "running" | "completed" | "error";
    message?: string;
    current_step: number;
    total_steps: number;
    current_page_number: number;
    page_timings: { page_number: number; elapsed_seconds: number }[];
    result_id?: string;
    error?: string;
    result?: {
        extracted_html: string;
        extracted_text: string;
        page_htmls: string[];
        page_texts: string[];
        page_timings: { page_number: number; elapsed_seconds: number }[];
        elapsed_seconds: number;
    };
};

export type ExcelSheet = {
    name: string;
    rows: string[][];
};
