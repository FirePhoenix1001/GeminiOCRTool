// word.js - Word Document Generation & Layout Refinement Module

import * as docx from 'docx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { readFileAsArrayBuffer } from './pdf.js';

// Parse subscript/superscript tags ^{} and _{} from plain text
export function parseSubscriptSuperscript(text) {
    const segments = [];
    let i = 0;
    const n = text.length;

    while (i < n) {
        const nextSuper = text.indexOf('^{', i);
        const nextSub = text.indexOf('_{', i);

        let startIdx = -1;
        let isSuper = false;

        if (nextSuper !== -1 && (nextSub === -1 || nextSuper < nextSub)) {
            startIdx = nextSuper;
            isSuper = true;
        } else if (nextSub !== -1 && (nextSuper === -1 || nextSub < nextSuper)) {
            startIdx = nextSub;
            isSuper = false;
        } else {
            segments.push({ text: text.substring(i), style: 'normal' });
            break;
        }

        if (startIdx > i) {
            segments.push({ text: text.substring(i, startIdx), style: 'normal' });
        }

        const closeIdx = text.indexOf('}', startIdx + 2);
        if (closeIdx !== -1) {
            const content = text.substring(startIdx + 2, closeIdx);
            segments.push({ text: content, style: isSuper ? 'super' : 'sub' });
            i = closeIdx + 1;
        } else {
            segments.push({ text: text.substring(startIdx), style: 'normal' });
            break;
        }
    }

    return segments;
}

// Build a formatted Word (.docx) document from OCR/detailed description texts
export async function buildWordDocument(text, filename) {
    const mathFunctions = ['sin', 'cos', 'tan', 'log'];
    const paragraphs = [];

    const lines = text.split('\n');
    for (let line of lines) {
        const runs = [];
        const segments = parseSubscriptSuperscript(line);

        for (let segment of segments) {
            const segmentText = segment.text;
            const style = segment.style; // 'normal', 'super', 'sub'

            // Split alphabetical variables to apply italics style
            const parts = segmentText.split(/([a-zA-Z]+)/);
            for (let part of parts) {
                if (!part) continue;

                const isAlphaLower = /^[a-z]+$/.test(part);
                const isMathFunc = mathFunctions.includes(part);

                const runOptions = {
                    text: part,
                    font: 'Times New Roman'
                };

                // Apply script formatting
                if (style === 'super') {
                    runOptions.superScript = true;
                } else if (style === 'sub') {
                    runOptions.subScript = true;
                }

                // Apply italic styles to standalone math variables (lower case, non-function words)
                if (isAlphaLower && !isMathFunc) {
                    runOptions.italic = true;
                }

                runs.push(new docx.TextRun(runOptions));
            }
        }

        paragraphs.push(new docx.Paragraph({
            children: runs
        }));
    }

    const doc = new docx.Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: 'Times New Roman'
                    }
                }
            }
        },
        sections: [{
            properties: {},
            children: paragraphs
        }]
    });

    const blob = await docx.Packer.toBlob(doc);
    saveAs(blob, `${filename}.docx`);
}

// Helper to trigger backup raw text downloads (.txt format)
export function triggerTxtDownload(content, filename) {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    saveAs(blob, filename);
}

// Optimize existing Word document file XML structure directly using JSZip
export async function optimizeWordFile(file, options, logMsg = console.log) {
    const { minusToHyphen, hyphenToMinus, convertSuper, convertSub } = options;
    
    // Read file via JSZip
    const arrayBuffer = await readFileAsArrayBuffer(file);
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // Fetch document XML
    let docXml = await zip.file('word/document.xml').async('text');

    // 1. Hyphen replacements inside XML text runs (<w:t>)
    if (hyphenToMinus) {
        docXml = docXml.replaceAll('-', '−');
    }
    if (minusToHyphen) {
        docXml = docXml.replaceAll('−', '-');
    }

    // 2. Parse superscripts ^{...} in Word XML format
    if (convertSuper) {
        let changed = true;
        while (changed) {
            const originalXml = docXml;
            docXml = docXml.replace(
                /<w:r>([\s\S]*?)<w:t>([\s\S]*?)\^\{([^\}]+)\}([\s\S]*?)<\/w:t><\/w:r>/g,
                (match, rPrContent, beforeText, superText, afterText) => {
                    const superRPr = rPrContent.includes('<w:rPr>') 
                        ? rPrContent.replace('<w:rPr>', '<w:rPr><w:vertAlign w:val="superscript"/>')
                        : `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>${rPrContent}`;
                    
                    return `<w:r>${rPrContent}<w:t>${beforeText}</w:t></w:r>` + 
                           `<w:r>${superRPr}<w:t>${superText}</w:t></w:r>` + 
                           `<w:r>${rPrContent}<w:t>${afterText}</w:t></w:r>`;
                }
            );
            if (docXml === originalXml) changed = false;
        }
    }

    // 3. Parse subscripts _{...} in Word XML format
    if (convertSub) {
        let changed = true;
        while (changed) {
            const originalXml = docXml;
            docXml = docXml.replace(
                /<w:r>([\s\S]*?)<w:t>([\s\S]*?)_\{([^\}]+)\}([\s\S]*?)<\/w:t><\/w:r>/g,
                (match, rPrContent, beforeText, subText, afterText) => {
                    const subRPr = rPrContent.includes('<w:rPr>') 
                        ? rPrContent.replace('<w:rPr>', '<w:rPr><w:vertAlign w:val="subscript"/>')
                        : `<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>${rPrContent}`;
                    
                    return `<w:r>${rPrContent}<w:t>${beforeText}</w:t></w:r>` + 
                           `<w:r>${subRPr}<w:t>${subText}</w:t></w:r>` + 
                           `<w:r>${rPrContent}<w:t>${afterText}</w:t></w:r>`;
                }
            );
            if (docXml === originalXml) changed = false;
        }
    }

    // Write modified XML back to zip structure
    zip.file('word/document.xml', docXml);

    // Build optimized zip blob and save it
    const optimizedBlob = await zip.generateAsync({ type: 'blob' });
    const originalStem = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
    saveAs(optimizedBlob, `${originalStem}_排版優化.docx`);
    logMsg(`✅ 優化完成，已觸發下載: ${originalStem}_排版優化.docx`);
}
