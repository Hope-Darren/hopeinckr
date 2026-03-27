import FileUploadZone from "./components/FileUploadZone";

export default function Home() {
  return (
    <div>
      <h1 className="text-2xl font-bold text-gray-800 mb-2">엑셀 파일 업로드</h1>
      <p className="text-sm text-gray-500 mb-6">
        .xlsx 또는 .xlsm 파일을 드래그하거나 클릭하여 업로드하세요.
      </p>
      <FileUploadZone />
    </div>
  );
}
