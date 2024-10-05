import fs, { promises as fsPromises } from "fs";
import pdfParse from "pdf-parse";
import { read, utils } from "xlsx";
import csvParse from "csv-parser";
import textract from "textract";
import MarkdownIt from "markdown-it";
import { getMimeType } from "./utils.js";
import https from 'https';
import http from 'http';
import { Buffer } from 'buffer';


// Инициализация Markdown парсера
const mdParser = new MarkdownIt();

export async function fromBuffer(
    buffer: Buffer,
    mimeType: string,
): Promise<string | void> {
    switch (mimeType) {
        case "text/plain":
            return buffer.toString("utf8");
        case "text/markdown":
            return mdParser.render(buffer.toString("utf8"));
        case "application/pdf":
            const pdfData = await pdfParse(buffer);
            return pdfData.text;
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        case "application/msword":
            return new Promise((resolve, reject) => {
                textract.fromBufferWithMime(mimeType, buffer, (error, text) => {
                    if (error) reject("Ошибка при извлечении текста из DOC/DOCX файла.");
                    else resolve(text);
                });
            });
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        case "application/vnd.ms-excel":
            try {
                const workbook = read(buffer, { type: "buffer" });
                let text = "";
                workbook.SheetNames.forEach((name) => {
                    const sheet = workbook.Sheets[name];
                    const data = utils.sheet_to_csv(sheet);
                    text += data;
                });
                return text;
            } catch (error) {
                return "Ошибка при обработке XLS/XLSX файла.";
            }
        case "text/csv":
            return new Promise((resolve, reject) => {
                let csvText = "";
                const stream = fs.createReadStream(buffer).pipe(csvParse());
                stream.on(
                    "data",
                    (data) => (csvText += Object.values(data).join(", ") + "\n"),
                );
                stream.on("end", () => resolve(csvText));
                stream.on("error", () => reject("Ошибка при чтении CSV файла."));
            });
        default:
            throw new Error("Формат файла не поддерживается.");
    }
}

export async function fromFile (
    filePath: string,
): Promise<string | void>  {
    const mime = getMimeType(filePath);
    const buffer = await fsPromises.readFile(filePath);

    if (!mime) {
        console.log("Unsupported filetype", filePath);
        return;
    }

    return fromBuffer(buffer, mime);
};


export async function fromUrl(url: string): Promise<string | void> {
    const {mime, buffer} = await fetchDocument(url);
    return fromBuffer(buffer, mime);
}


interface DocumentResponse {
    mime: string;
    buffer: Buffer;
}

function fetchDocument(url: string): Promise<DocumentResponse> {
    return new Promise((resolve, reject) => {
        const client = url.startsWith('https') ? https : http;

        client.get(url, (response) => {
            if (response.statusCode !== 200) {
                reject(new Error(`Ошибка при получении документа: ${response.statusCode}`));
                return;
            }

            const mime = response.headers['content-type'] || getMimeType(url) || null; // Получаем MIME-тип

            if (!mime) {
                reject(new Error('Unknown mime type'));
                return
            }

            const chunks: Buffer[] = [];

            // Читаем данные по частям и сохраняем их в массиве `chunks`
            response.on('data', (chunk: Buffer) => {
                chunks.push(chunk);
            });

            // Когда документ полностью загружен, соединяем части в один буфер
            response.on('end', () => {
                const buffer = Buffer.concat(chunks); // Собираем буфер из всех частей
                resolve({ mime, buffer });
            });

            // Обрабатываем ошибки
            response.on('error', (err) => {
                reject(err);
            });
        }).on('error', (err) => {
            reject(err);
        });
    });
}
