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
                    font: {
                        ascii: 'Times New Roman',
                        eastAsia: '標楷體'
                    }
                };

                // Apply script formatting
                if (style === 'super') {
                    runOptions.superScript = true;
                } else if (style === 'sub') {
                    runOptions.subScript = true;
                }

                // Apply italic styles to standalone math variables (lower case, non-function words)
                if (isAlphaLower && !isMathFunc) {
                    runOptions.italics = true;
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
                        font: {
                            ascii: 'Times New Roman',
                            eastAsia: '標楷體'
                        }
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

const RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 
    'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 
    'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 
    'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 
    'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 
    'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'
];

function insertRPrChild(rPr, newElem, xmlDoc) {
    const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const localName = newElem.localName;
    const targetIdx = RPR_ORDER.indexOf(localName);
    
    if (targetIdx === -1) {
        rPr.appendChild(newElem);
        return;
    }
    
    const existing = rPr.getElementsByTagNameNS(NS_W, localName)[0];
    if (existing) {
        rPr.replaceChild(newElem, existing);
        return;
    }
    
    const children = Array.from(rPr.childNodes).filter(node => node.nodeType === 1);
    for (let child of children) {
        const childLocalName = child.localName;
        const childIdx = RPR_ORDER.indexOf(childLocalName);
        if (childIdx !== -1 && childIdx > targetIdx) {
            rPr.insertBefore(newElem, child);
            return;
        }
    }
    
    rPr.appendChild(newElem);
}

// Optimize existing Word document file XML structure directly using JSZip
export async function optimizeWordFile(file, options, logMsg = console.log) {
    const { minusToHyphen, hyphenToMinus, convertSuper, convertSub, convertFont, convertItalic } = options;
    
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

    // 4. Font and Italic conversions in Word XML format
    if (convertFont || convertItalic) {
        logMsg(`⚙️ 正在進行字型與斜體樣式處理...`);
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(docXml, "application/xml");
            const parserError = xmlDoc.getElementsByTagName("parsererror")[0];
            if (parserError) {
                logMsg(`⚠️ Word XML 解析時包含錯誤，跳過字型與斜體最佳化。`);
            } else {
                const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                const mathFunctions = ['sin', 'cos', 'tan', 'log'];
                
                const paragraphs = xmlDoc.getElementsByTagNameNS(NS_W, 'p');
                for (let i = 0; i < paragraphs.length; i++) {
                    const p = paragraphs[i];
                    const runs = Array.from(p.getElementsByTagNameNS(NS_W, 'r'));
                    
                    for (let rNode of runs) {
                        const tNode = rNode.getElementsByTagNameNS(NS_W, 't')[0];
                        if (!tNode || !tNode.textContent) {
                            continue;
                        }
                        
                        const oldText = tNode.textContent;
                        const parts = oldText.split(/([a-zA-Z]+)/);
                        
                        // If there is only one part and no splitting is needed
                        if (parts.length <= 1) {
                            let rPr = rNode.getElementsByTagNameNS(NS_W, 'rPr')[0];
                            if (!rPr) {
                                rPr = xmlDoc.createElementNS(NS_W, 'w:rPr');
                                rNode.insertBefore(rPr, tNode);
                            }
                            
                            if (convertFont) {
                                let rFonts = rPr.getElementsByTagNameNS(NS_W, 'rFonts')[0];
                                if (!rFonts) {
                                    rFonts = xmlDoc.createElementNS(NS_W, 'w:rFonts');
                                }
                                rFonts.setAttributeNS(NS_W, 'w:ascii', 'Times New Roman');
                                rFonts.setAttributeNS(NS_W, 'w:hAnsi', 'Times New Roman');
                                rFonts.setAttributeNS(NS_W, 'w:eastAsia', '標楷體');
                                insertRPrChild(rPr, rFonts, xmlDoc);
                            }
                            
                            if (convertItalic) {
                                const isAlphaLower = /^[a-z]+$/.test(oldText);
                                const isMathFunc = mathFunctions.includes(oldText);
                                const shouldItalic = isAlphaLower && !isMathFunc;
                                
                                if (shouldItalic) {
                                    let iElem = rPr.getElementsByTagNameNS(NS_W, 'i')[0];
                                    if (!iElem) {
                                        iElem = xmlDoc.createElementNS(NS_W, 'w:i');
                                        insertRPrChild(rPr, iElem, xmlDoc);
                                    }
                                    let iCsElem = rPr.getElementsByTagNameNS(NS_W, 'iCs')[0];
                                    if (!iCsElem) {
                                        iCsElem = xmlDoc.createElementNS(NS_W, 'w:iCs');
                                        insertRPrChild(rPr, iCsElem, xmlDoc);
                                    }
                                } else {
                                    let iElem = rPr.getElementsByTagNameNS(NS_W, 'i')[0];
                                    if (iElem) {
                                        rPr.removeChild(iElem);
                                    }
                                    let iCsElem = rPr.getElementsByTagNameNS(NS_W, 'iCs')[0];
                                    if (iCsElem) {
                                        rPr.removeChild(iCsElem);
                                    }
                                }
                            }
                            continue;
                        }
                        
                        let currentRefRun = rNode;
                        let firstPart = true;
                        
                        for (let part of parts) {
                            if (!part) continue;
                            
                            let targetRun;
                            if (firstPart) {
                                targetRun = rNode;
                                const tElem = targetRun.getElementsByTagNameNS(NS_W, 't')[0];
                                if (tElem) {
                                    tElem.textContent = part;
                                    tElem.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
                                }
                                firstPart = false;
                            } else {
                                targetRun = rNode.cloneNode(true);
                                const tElem = targetRun.getElementsByTagNameNS(NS_W, 't')[0];
                                if (tElem) {
                                    tElem.textContent = part;
                                    tElem.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
                                }
                                
                                const parent = currentRefRun.parentNode;
                                if (currentRefRun.nextSibling) {
                                    parent.insertBefore(targetRun, currentRefRun.nextSibling);
                                } else {
                                    parent.appendChild(targetRun);
                                }
                                currentRefRun = targetRun;
                            }
                            
                            let targetRPr = targetRun.getElementsByTagNameNS(NS_W, 'rPr')[0];
                            if (!targetRPr) {
                                targetRPr = xmlDoc.createElementNS(NS_W, 'w:rPr');
                                const tElem = targetRun.getElementsByTagNameNS(NS_W, 't')[0];
                                if (tElem) {
                                    targetRun.insertBefore(targetRPr, tElem);
                                } else {
                                    targetRun.appendChild(targetRPr);
                                }
                            }
                            
                            if (convertFont) {
                                let rFonts = targetRPr.getElementsByTagNameNS(NS_W, 'rFonts')[0];
                                if (!rFonts) {
                                    rFonts = xmlDoc.createElementNS(NS_W, 'w:rFonts');
                                }
                                rFonts.setAttributeNS(NS_W, 'w:ascii', 'Times New Roman');
                                rFonts.setAttributeNS(NS_W, 'w:hAnsi', 'Times New Roman');
                                rFonts.setAttributeNS(NS_W, 'w:eastAsia', '標楷體');
                                insertRPrChild(targetRPr, rFonts, xmlDoc);
                            }
                            
                            if (convertItalic) {
                                const isAlphaLower = /^[a-z]+$/.test(part);
                                const isMathFunc = mathFunctions.includes(part);
                                const shouldItalic = isAlphaLower && !isMathFunc;
                                
                                if (shouldItalic) {
                                    let iElem = targetRPr.getElementsByTagNameNS(NS_W, 'i')[0];
                                    if (!iElem) {
                                        iElem = xmlDoc.createElementNS(NS_W, 'w:i');
                                        insertRPrChild(targetRPr, iElem, xmlDoc);
                                    }
                                    let iCsElem = targetRPr.getElementsByTagNameNS(NS_W, 'iCs')[0];
                                    if (!iCsElem) {
                                        iCsElem = xmlDoc.createElementNS(NS_W, 'w:iCs');
                                        insertRPrChild(targetRPr, iCsElem, xmlDoc);
                                    }
                                } else {
                                    let iElem = targetRPr.getElementsByTagNameNS(NS_W, 'i')[0];
                                    if (iElem) {
                                        targetRPr.removeChild(iElem);
                                    }
                                    let iCsElem = targetRPr.getElementsByTagNameNS(NS_W, 'iCs')[0];
                                    if (iCsElem) {
                                        targetRPr.removeChild(iCsElem);
                                    }
                                }
                            }
                        }
                    }
                }
                
                const serializer = new XMLSerializer();
                docXml = serializer.serializeToString(xmlDoc);
            }
        } catch (e) {
            logMsg(`⚠️ 樣式優化處理失敗，跳過字型/斜體最佳化: ${e.message}`);
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
