
import { PDFDocument } from "pdf-lib";
import html2canvas from "@html2canvas/html2canvas";
import * as mammoth from "mammoth/mammoth.browser";
import * as XLSX from "xlsx";
import { PptxRenderer } from "pptx-browser";

const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const convertBtn = document.getElementById("convertBtn");
const status = document.getElementById("status");

let selectedFile = null;

const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;



fileInput.addEventListener("change", () => {
    selectedFile = fileInput.files[0];

    if (!selectedFile) return;

    fileInfo.textContent =
        `${selectedFile.name} (${formatSize(selectedFile.size)})`;

    convertBtn.disabled = false;
    status.textContent = "Ready to convert";
});

function formatSize(bytes) {
    return (bytes / 1024 / 1024).toFixed(2) + " MB";
}



convertBtn.addEventListener("click", async () => {

    if (!selectedFile) return;

    convertBtn.disabled = true;
    status.textContent = "Converting...";

    try {

        const extension = selectedFile.name
            .split(".")
            .pop()
            .toLowerCase();

        let pdfBytes;

        switch (extension) {

            case "jpg":
            case "jpeg":
            case "png":
            case "webp":
                pdfBytes = await convertImage(selectedFile);
                break;

            case "txt":
                pdfBytes = await convertText(selectedFile);
                break;

            case "html":
            case "htm":
                pdfBytes = await convertHTML(selectedFile);
                break;

            case "docx":
                pdfBytes = await convertDOCX(selectedFile);
                break;

            case "xlsx":
                pdfBytes = await convertXLSX(selectedFile);
                break;

            case "pptx":
                pdfBytes = await convertPPTX(selectedFile);
                break;

            default:
                throw new Error("Unsupported file format");
        }

        downloadPDF(pdfBytes, selectedFile.name);

        status.textContent = "Conversion completed!";

    } catch (error) {

        console.error(error);

        status.textContent =
            "Error: " + error.message;

    } finally {

        convertBtn.disabled = false;

    }
});

// ------------------------------
// 3. Image to PDF
// ------------------------------

async function convertImage(file) {

    const pdf = await PDFDocument.create();

    const bitmap = await createImageBitmap(file);

    const canvas = document.createElement("canvas");

    canvas.width = bitmap.width;
    canvas.height = bitmap.height;

    const context = canvas.getContext("2d");

    context.drawImage(bitmap, 0, 0);

    const imageBytes = await new Promise((resolve, reject) => {

        canvas.toBlob(blob => {

            if (!blob) {
                reject(new Error("Image conversion failed"));
                return;
            }

            blob.arrayBuffer().then(resolve);

        }, "image/png");

    });

    const image = await pdf.embedPng(imageBytes);

    const page = pdf.addPage([A4_WIDTH, A4_HEIGHT]);

    const margin = 25;

    const scale = Math.min(
        (A4_WIDTH - margin * 2) / image.width,
        (A4_HEIGHT - margin * 2) / image.height
    );

    const width = image.width * scale;
    const height = image.height * scale;

    page.drawImage(image, {
        x: (A4_WIDTH - width) / 2,
        y: (A4_HEIGHT - height) / 2,
        width,
        height
    });

    bitmap.close();

    return await pdf.save();
}

// ------------------------------
// 4. HTML to PDF
// ------------------------------

async function convertHTML(file) {

    const html = await file.text();

    return renderHTMLToPDF(html);
}



async function convertText(file) {

    const text = await file.text();

    const escaped = text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    const html = `
        <h2>${escapeText(file.name)}</h2>
        <pre class="plain-text">${escaped}</pre>
    `;

    return renderHTMLToPDF(html);
}

function escapeText(text) {
    return text.replace(/[&<>"']/g, char => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
    })[char]);
}



async function convertDOCX(file) {

    const arrayBuffer = await file.arrayBuffer();

    const result = await mammoth.convertToHtml({
        arrayBuffer
    });

    if (!result.value.trim()) {
        throw new Error("No document content found");
    }

    return renderHTMLToPDF(result.value);
}



async function convertXLSX(file) {

    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer);

    let html = `<h2>${escapeText(file.name)}</h2>`;

    for (const sheetName of workbook.SheetNames) {

        const sheet = workbook.Sheets[sheetName];

        html += `<h2>${escapeText(sheetName)}</h2>`;

        html += XLSX.utils.sheet_to_html(sheet);

    }

    return renderHTMLToPDF(html);
}



async function convertPPTX(file) {

    const renderer = new PptxRenderer();

    await renderer.load(file);

    const pdf = await PDFDocument.create();

    const slideCanvas = document.createElement("canvas");

    slideCanvas.width = 1280;
    slideCanvas.height = 720;

    for (let i = 0; i < renderer.slideCount; i++) {

        await renderer.renderSlide(i, slideCanvas, 1280);

        const imageBytes = await new Promise((resolve, reject) => {

            slideCanvas.toBlob(blob => {

                if (!blob) {
                    reject(new Error("Slide rendering failed"));
                    return;
                }

                blob.arrayBuffer().then(resolve);

            }, "image/png");

        });

        const image = await pdf.embedPng(imageBytes);

        const page = pdf.addPage([A4_WIDTH, A4_HEIGHT]);

        const margin = 20;

        const scale = Math.min(
            (A4_WIDTH - margin * 2) / image.width,
            (A4_HEIGHT - margin * 2) / image.height
        );

        const width = image.width * scale;
        const height = image.height * scale;

        page.drawImage(image, {
            x: (A4_WIDTH - width) / 2,
            y: (A4_HEIGHT - height) / 2,
            width,
            height
        });

    }

    return await pdf.save();
}


async function renderHTMLToPDF(html) {

    const host = document.createElement("div");

    host.style.cssText = `
        position: fixed;
        left: -10000px;
        top: 0;
        width: 794px;
        padding: 35px;
        background: white;
        color: black;
        font-family: Arial, sans-serif;
        font-size: 14px;
        line-height: 1.5;
        overflow-wrap: anywhere;
    `;

    const style = document.createElement("style");

    style.textContent = `
        * {
            box-sizing: border-box;
        }

        h1, h2, h3 {
            overflow-wrap: anywhere;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            margin-bottom: 20px;
        }

        td, th {
            border: 1px solid #999;
            padding: 5px;
        }

        img {
            max-width: 100%;
            height: auto;
        }

        pre {
            white-space: pre-wrap;
            overflow-wrap: anywhere;
            font-family: monospace;
        }

        .plain-text {
            white-space: pre-wrap;
            overflow-wrap: anywhere;
        }
    `;

    host.appendChild(style);

    // Parse input as HTML and copy only safe elements.
    const parsed = new DOMParser().parseFromString(
        html,
        "text/html"
    );

    const allowed = new Set([
        "P", "DIV", "SPAN", "BR",
        "H1", "H2", "H3", "H4", "H5", "H6",
        "UL", "OL", "LI",
        "TABLE", "THEAD", "TBODY", "TR", "TD", "TH",
        "STRONG", "B", "EM", "I", "U",
        "PRE", "BLOCKQUOTE", "HR", "IMG"
    ]);

    function copySafe(node, parent) {

        if (node.nodeType === Node.TEXT_NODE) {
            parent.appendChild(
                document.createTextNode(node.textContent)
            );
            return;
        }

        if (node.nodeType !== Node.ELEMENT_NODE) return;

        const tag = node.tagName;

        if (!allowed.has(tag)) {

            for (const child of node.childNodes) {
                copySafe(child, parent);
            }

            return;
        }

        const element = document.createElement(
            tag.toLowerCase()
        );

        if (tag === "IMG") {

            const src = node.getAttribute("src") || "";

            if (!src.startsWith("data:image/")) return;

            element.setAttribute("src", src);
            element.setAttribute("alt", "Image");

        }

        parent.appendChild(element);

        for (const child of node.childNodes) {
            copySafe(child, element);
        }
    }

    for (const child of parsed.body.childNodes) {
        copySafe(child, host);
    }

    document.body.appendChild(host);

    try {

        const canvas = await html2canvas(host, {
            scale: 1.5,
            backgroundColor: "#ffffff",
            logging: false,
            width: 794,
            windowWidth: 794
        });

        const pdf = await PDFDocument.create();

        const pageHeight = Math.floor(
            canvas.width * A4_HEIGHT / A4_WIDTH
        );

        for (
            let y = 0;
            y < canvas.height;
            y += pageHeight
        ) {

            const sliceHeight = Math.min(
                pageHeight,
                canvas.height - y
            );

            const pageCanvas = document.createElement("canvas");

            pageCanvas.width = canvas.width;
            pageCanvas.height = sliceHeight;

            const ctx = pageCanvas.getContext("2d");

            ctx.drawImage(
                canvas,
                0, y,
                canvas.width, sliceHeight,
                0, 0,
                canvas.width, sliceHeight
            );

            const bytes = await new Promise((resolve, reject) => {

                pageCanvas.toBlob(blob => {

                    if (!blob) {
                        reject(new Error("PDF page rendering failed"));
                        return;
                    }

                    blob.arrayBuffer().then(resolve);

                }, "image/png");

            });

            const image = await pdf.embedPng(bytes);

            const page = pdf.addPage([
                A4_WIDTH,
                A4_HEIGHT
            ]);

            const renderedHeight =
                A4_WIDTH * sliceHeight / canvas.width;

            page.drawImage(image, {
                x: 0,
                y: A4_HEIGHT - renderedHeight,
                width: A4_WIDTH,
                height: renderedHeight
            });

        }

        canvas.width = 0;
        canvas.height = 0;

        return await pdf.save();

    } finally {

        host.remove();

    }
}



function downloadPDF(bytes, originalName) {

    const blob = new Blob(
        [bytes],
        { type: "application/pdf" }
    );

    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");

    const baseName = originalName.replace(/\.[^.]+$/, "");

    link.href = url;
    link.download = baseName + ".pdf";

    document.body.appendChild(link);

    link.click();

    link.remove();

    setTimeout(() => URL.revokeObjectURL(url), 60000);
}
