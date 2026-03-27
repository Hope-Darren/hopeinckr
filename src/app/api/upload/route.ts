import { NextRequest, NextResponse } from "next/server";
import formidable from "formidable";
import fs from "fs";
import path from "path";
import { Readable } from "stream";
import { IncomingMessage } from "http";
import { v4 as uuidv4 } from "uuid";

const ALLOWED_EXTENSIONS = new Set([".xlsx", ".xlsm"]);
const ALLOWED_MIME_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel.sheet.macroenabled.12",
  // Some browsers send generic mime type for .xlsm
  "application/vnd.ms-excel",
]);

export async function POST(req: NextRequest) {
  const uploadsDir = path.join(process.cwd(), "uploads");
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  // Convert Web Request to Node.js IncomingMessage-compatible stream
  const arrayBuffer = await req.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  const readable = Readable.from(buffer) as unknown as IncomingMessage;
  readable.headers = Object.fromEntries(req.headers.entries());
  readable.method = "POST";

  const form = formidable({
    uploadDir: uploadsDir,
    keepExtensions: true,
    maxFileSize: 50 * 1024 * 1024, // 50MB
    filename: (_name, ext) => `${uuidv4()}${ext}`,
  });

  return new Promise<NextResponse>((resolve) => {
    form.parse(readable, (err, _fields, files) => {
      if (err) {
        resolve(
          NextResponse.json({ success: false, error: "파싱 오류: " + err.message }, { status: 400 })
        );
        return;
      }

      const fileField = files.file;
      if (!fileField) {
        resolve(
          NextResponse.json({ success: false, error: "파일이 없습니다." }, { status: 400 })
        );
        return;
      }

      const uploaded = Array.isArray(fileField) ? fileField : [fileField];
      const results: Array<{ originalName: string; savedName: string; size: number }> = [];

      for (const file of uploaded) {
        const originalName = file.originalFilename ?? "unknown";
        const ext = path.extname(originalName).toLowerCase();
        const mime = file.mimetype ?? "";

        if (!ALLOWED_EXTENSIONS.has(ext) && !ALLOWED_MIME_TYPES.has(mime)) {
          // Remove the temp file
          fs.unlink(file.filepath, () => {});
          resolve(
            NextResponse.json(
              { success: false, error: `.xlsx 또는 .xlsm 파일만 허용됩니다. (받은 파일: ${originalName})` },
              { status: 422 }
            )
          );
          return;
        }

        results.push({
          originalName,
          savedName: path.basename(file.filepath),
          size: file.size,
        });
      }

      resolve(NextResponse.json({ success: true, files: results }));
    });
  });
}
