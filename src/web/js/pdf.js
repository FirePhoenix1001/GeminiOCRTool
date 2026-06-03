// pdf.js - PDF Page Extraction and Reading Module

// Configure PDF.js worker
if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
} else {
    // Retry worker binding if library loads late
    document.addEventListener('DOMContentLoaded', () => {
        if (window.pdfjsLib) {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    });
}

// Helper to read file as ArrayBuffer
export function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = (e) => reject(e);
        reader.readAsArrayBuffer(file);
    });
}

// Helper to read image file directly as Base64
export function readImageAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = (e) => reject(e);
        reader.readAsDataURL(file);
    });
}

// Convert PDF page to PNG base64 representation
export async function convertPdfPageToPngBase64(file, pageNum) {
    const pdfjsLib = window.pdfjsLib;
    if (!pdfjsLib) {
        throw new Error('PDF.js 函式庫尚未載入，請確認網路連線是否正常。');
    }
    
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const page = await pdf.getPage(pageNum);

    // Render page to canvas with high resolution scale
    const scale = 2.0; // Higher DPI for OCR accuracy
    const viewport = page.getViewport({ scale: scale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    const renderContext = {
        canvasContext: context,
        viewport: viewport
    };

    await page.render(renderContext).promise;

    const dataUrl = canvas.toDataURL('image/png');
    // Extract base64 part
    return dataUrl.split(',')[1];
}

// Function to fetch page count of a PDF file
export async function getPdfPageCount(file) {
    const pdfjsLib = window.pdfjsLib;
    if (!pdfjsLib) {
        throw new Error('PDF.js 函式庫尚未載入。');
    }
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    return pdf.numPages;
}
