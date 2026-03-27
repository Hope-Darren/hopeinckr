"use client";

import { useCallback, useState } from "react";
import { useDropzone } from "react-dropzone";
import { v4 as uuidv4 } from "uuid";
import FileList from "./FileList";

export type UploadFile = {
  id: string;
  file: File;
  status: "pending" | "uploading" | "done" | "error";
  progress: number;
  error?: string;
};

const ACCEPTED_TYPES = {
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
  "application/vnd.ms-excel.sheet.macroenabled.12": [".xlsm"],
};

export default function FileUploadZone() {
  const [files, setFiles] = useState<UploadFile[]>([]);
  const [rejectedMessage, setRejectedMessage] = useState<string | null>(null);

  const updateFile = useCallback(
    (id: string, patch: Partial<Omit<UploadFile, "id" | "file">>) => {
      setFiles((prev) =>
        prev.map((f) => (f.id === id ? { ...f, ...patch } : f))
      );
    },
    []
  );

  const uploadFile = useCallback(
    (item: UploadFile) => {
      updateFile(item.id, { status: "uploading", progress: 0 });

      const formData = new FormData();
      formData.append("file", item.file);

      const xhr = new XMLHttpRequest();

      xhr.upload.onprogress = (e) => {
        if (e.lengthComputable) {
          const pct = Math.round((e.loaded / e.total) * 100);
          updateFile(item.id, { progress: pct });
        }
      };

      xhr.onload = () => {
        if (xhr.status === 200) {
          updateFile(item.id, { status: "done", progress: 100 });
        } else {
          let msg = "업로드 실패";
          try {
            const res = JSON.parse(xhr.responseText);
            if (res.error) msg = res.error;
          } catch {}
          updateFile(item.id, { status: "error", error: msg });
        }
      };

      xhr.onerror = () => {
        updateFile(item.id, { status: "error", error: "네트워크 오류" });
      };

      xhr.open("POST", "/api/upload");
      xhr.send(formData);
    },
    [updateFile]
  );

  const onDropAccepted = useCallback(
    (accepted: File[]) => {
      setRejectedMessage(null);
      const newItems: UploadFile[] = accepted.map((file) => ({
        id: uuidv4(),
        file,
        status: "pending",
        progress: 0,
      }));
      setFiles((prev) => [...prev, ...newItems]);
      newItems.forEach((item) => uploadFile(item));
    },
    [uploadFile]
  );

  const onDropRejected = useCallback(() => {
    setRejectedMessage(".xlsx 또는 .xlsm 파일만 업로드할 수 있습니다.");
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDropAccepted,
    onDropRejected,
    accept: ACCEPTED_TYPES,
    multiple: true,
  });

  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  };

  return (
    <div>
      <div
        {...getRootProps()}
        className={`border-2 border-dashed rounded-xl p-10 text-center cursor-pointer transition-colors ${
          isDragActive
            ? "border-blue-500 bg-blue-50"
            : "border-gray-300 bg-white hover:border-blue-400 hover:bg-blue-50"
        }`}
      >
        <input {...getInputProps()} />
        <svg
          className="mx-auto mb-4 w-12 h-12 text-gray-400"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeWidth={1.5}
            d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
          />
        </svg>
        {isDragActive ? (
          <p className="text-blue-600 font-medium">여기에 놓으세요...</p>
        ) : (
          <>
            <p className="text-gray-600 font-medium">
              엑셀 파일을 여기로 드래그하거나
            </p>
            <p className="text-blue-500 mt-1 text-sm">클릭하여 파일 선택</p>
          </>
        )}
        <p className="text-gray-400 text-xs mt-3">지원 형식: .xlsx, .xlsm</p>
      </div>

      {rejectedMessage && (
        <p className="mt-2 text-sm text-red-500">{rejectedMessage}</p>
      )}

      <FileList files={files} onRemove={removeFile} />
    </div>
  );
}
