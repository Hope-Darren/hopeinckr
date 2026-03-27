import ProgressBar from "./ProgressBar";
import { UploadFile } from "./FileUploadZone";

interface FileListProps {
  files: UploadFile[];
  onRemove: (id: string) => void;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const statusLabel: Record<UploadFile["status"], string> = {
  pending: "대기 중",
  uploading: "업로드 중...",
  done: "완료",
  error: "오류",
};

const statusColor: Record<UploadFile["status"], string> = {
  pending: "text-gray-500",
  uploading: "text-blue-600",
  done: "text-green-600",
  error: "text-red-600",
};

export default function FileList({ files, onRemove }: FileListProps) {
  if (files.length === 0) return null;

  return (
    <ul className="mt-4 space-y-3">
      {files.map((item) => (
        <li
          key={item.id}
          className="bg-white border border-gray-200 rounded-lg px-4 py-3 shadow-sm"
        >
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2 min-w-0">
              {/* Spreadsheet icon */}
              <svg
                className="w-6 h-6 text-green-600 shrink-0"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M9 17v-6m3 6v-3m3 3v-9M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H9L5 9v10a2 2 0 002 2z"
                />
              </svg>
              <div className="min-w-0">
                <p className="text-sm font-medium text-gray-800 truncate">
                  {item.file.name}
                </p>
                <p className="text-xs text-gray-400">{formatSize(item.file.size)}</p>
              </div>
            </div>
            <div className="flex items-center gap-3 ml-3 shrink-0">
              <span className={`text-xs font-medium ${statusColor[item.status]}`}>
                {item.error ? `오류: ${item.error}` : statusLabel[item.status]}
              </span>
              {(item.status === "pending" || item.status === "done" || item.status === "error") && (
                <button
                  onClick={() => onRemove(item.id)}
                  className="text-gray-300 hover:text-gray-500 transition-colors"
                  aria-label="파일 제거"
                >
                  <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
                    <path
                      fillRule="evenodd"
                      d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z"
                      clipRule="evenodd"
                    />
                  </svg>
                </button>
              )}
            </div>
          </div>
          <ProgressBar progress={item.progress} status={item.status} />
        </li>
      ))}
    </ul>
  );
}
