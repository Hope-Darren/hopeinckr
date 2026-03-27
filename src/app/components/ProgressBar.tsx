type Status = "pending" | "uploading" | "done" | "error";

interface ProgressBarProps {
  progress: number;
  status: Status;
}

export default function ProgressBar({ progress, status }: ProgressBarProps) {
  const colorClass =
    status === "done"
      ? "bg-green-500"
      : status === "error"
      ? "bg-red-500"
      : "bg-blue-500";

  return (
    <div className="w-full bg-gray-200 rounded-full h-2 mt-2">
      <div
        className={`h-2 rounded-full transition-all duration-300 ${colorClass}`}
        style={{ width: `${progress}%` }}
      />
    </div>
  );
}
