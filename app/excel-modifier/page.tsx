"use client";

import { useState, useEffect } from "react";
import {
  Upload,
  Download,
  CheckCircle2,
  AlertCircle,
  Heart,
  Loader2,
  Moon,
  Sun,
} from "lucide-react";


export default function ExcelTrafficModifier() {
  const [file, setFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [log, setLog] = useState<string[]>([]);
  const [percentage, setPercentage] = useState(13);
  const [operation, setOperation] = useState<"increase" | "decrease">(
    "increase"
  );
  const [dragActive, setDragActive] = useState(false);
  const [downloadFilename, setDownloadFilename] = useState(
    "Modified_Traffic_Counts.xlsx"
  );
  const [darkMode, setDarkMode] = useState(false);
  const [downloading, setDownloading] = useState(false);


  useEffect(() => {
    const saved = localStorage.getItem("darkMode");
    if (saved) {
      setDarkMode(saved === "true");
      document.documentElement.classList.toggle("dark", saved === "true");
    }
  }, []);

  const toggleDarkMode = () => {
    const newMode = !darkMode;
    setDarkMode(newMode);
    localStorage.setItem("darkMode", String(newMode));
    if (newMode) {
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }
  };

  const addLog = (git branchmessage: string) => {
    setLog((prev) => [
      ...prev,
      `${new Date().toLocaleTimeString()}: ${message}`,
    ]);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (uploadedFile) {
      setFile(uploadedFile);
      setDownloadUrl(null);
      setStatus("");
      setLog([]);
      addLog(`File selected: ${uploadedFile.name}`);
    }
  };

  const handlePercentageChange = (value: number) => {
    setPercentage(value);
    setDownloadUrl(null);
  };

  const handleOperationChange = (value: "increase" | "decrease") => {
    setOperation(value);
    setDownloadUrl(null);
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files?.[0]) {
      setFile(e.dataTransfer.files[0]);
      setDownloadUrl(null);
      setStatus("");
      setLog([]);
      addLog(`File selected: ${e.dataTransfer.files[0].name}`);
    }
  };

  const handleDownload = async () => {
    if (!downloadUrl) return;
    setDownloading(true);
    
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = downloadFilename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    setTimeout(() => setDownloading(false), 1000);
  };

  const processExcel = async () => {
    if (!file) {
      setStatus("Please select a file first");
      return;
    }

    setProcessing(true);
    setStatus("");
    setLog([]);
    addLog("Starting processing...");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("percentage", percentage.toString());
      formData.append("operation", operation);

      const response = await fetch("/api/modify-excel", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Processing failed");
      }

      const processLog = response.headers.get("X-Process-Log");
      if (processLog) {
        const decodedLog = atob(processLog);
        decodedLog.split("\n").forEach((line) => {
          if (line.trim()) addLog(line);
        });
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);

      const originalName = file.name;
      const nameWithoutExt = originalName.replace(/\.[^/.]+$/, "");
      const ext = originalName.match(/\.[^/.]+$/)?.[0] || ".xlsx";
      setDownloadFilename(`${nameWithoutExt}_mody${ext}`);

      setStatus("File processed successfully!");
      addLog("✓ Processing complete with formatting preserved!");
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : "Unknown error";
      console.error(error);
      setStatus(`Error: ${errorMessage}`);
      addLog(`✗ Error: ${errorMessage}`);
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-animated-gradient p-8 relative overflow-hidden">
      <div className="max-w-4xl mx-auto bg-white/95 dark:bg-gray-800/95 backdrop-blur-sm rounded-lg shadow-xl p-8 transition-colors duration-300 relative z-10">
        {/* Dark Mode Toggling */}
        <div className="flex justify-end mb-4">
          <button
            onClick={toggleDarkMode}
            className="p-2 rounded-lg bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600 transition-colors"
            aria-label="Toggle dark mode"
          >
            {darkMode ? (
              <Sun className="w-5 h-5 text-yellow-400" />
            ) : (
              <Moon className="w-5 h-5 text-gray-700 dark:text-gray-300" />
            )}
          </button>
        </div>

        <h1 className="text-3xl font-bold mb-2 text-gray-900 dark:text-white">
          Traffic Count Modifier
        </h1>
        <p className="text-gray-600 dark:text-gray-300 mb-6">Vehicular traffic modifications tool</p>

        {/* Upload with drag and drop */}
        <label className="block mb-6 cursor-pointer">
          <div
            className={`border-2 border-dashed rounded-lg p-6 transition-colors flex flex-col items-center ${
              dragActive
                ? "border-indigo-600 bg-indigo-50 dark:bg-indigo-900/20"
                : "border-gray-300 dark:border-gray-600 hover:border-indigo-500 dark:hover:border-indigo-400"
            }`}
            onDragEnter={handleDrag}
            onDragLeave={handleDrag}
            onDragOver={handleDrag}
            onDrop={handleDrop}
          >
            <Upload className="w-8 h-8 text-gray-400 dark:text-gray-500 mb-2" />
            <span className="text-gray-600 dark:text-gray-300">
              {file ? file.name : "Click or drag & drop Excel file"}
            </span>
          </div>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="hidden"
          />
        </label>

        {/* Configuration */}
        <div className="grid grid-cols-2 gap-4 mb-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 dark:text-gray-200 mb-2">
              Operation
            </label>
            <select
              value={operation}
              onChange={(e) =>
                handleOperationChange(e.target.value as "increase" | "decrease")
              }
              className="w-full px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
            >
              <option value="increase">Increase</option>
              <option value="decrease">Decrease</option>
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 dark:text-gray-200 mb-2">
              Percentage (%)
            </label>
            <input
              type="number"
              value={percentage}
              onChange={(e) => handlePercentageChange(Number(e.target.value))}
              onFocus={(e) => e.target.select()}
              min="1"
              max="100"
              className="w-full px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
            />
          </div>
        </div>

        {/* Process button */}
        <button
          onClick={processExcel}
          disabled={!file || processing}
          className="w-full bg-indigo-600 dark:bg-indigo-500 text-white py-3 px-6 rounded-lg font-medium hover:bg-indigo-700 dark:hover:bg-indigo-600 disabled:bg-gray-300 dark:disabled:bg-gray-600 disabled:cursor-not-allowed mb-6 flex items-center justify-center gap-2 transition-colors"
        >
          {processing ? (
            <>
              <Loader2 className="w-5 h-5 animate-spin" />
              Processing<span className="animate-pulse">...</span>
            </>
          ) : (
            `${
              operation === "increase" ? "Increase" : "Decrease"
            } by ${percentage}%`
          )}
        </button>
        {/* Status */}
        {status && !processing && (
          <div
            className={`p-4 rounded-lg mb-6 flex items-start gap-3 ${
              status.startsWith("Success")
                ? "bg-green-50 dark:bg-green-900/20 text-green-800 dark:text-green-300"
                : status.startsWith("Error")
                ? "bg-red-50 dark:bg-red-900/20 text-red-800 dark:text-red-300"
                : "bg-blue-50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300"
            }`}
          >
            {status.startsWith("Success") ? (
              <CheckCircle2 className="w-5 h-5 mt-0.5" />
            ) : status.startsWith("Error") ? (
              <AlertCircle className="w-5 h-5 mt-0.5" />
            ) : (
              <AlertCircle className="w-5 h-5 mt-0.5" />
            )}
            <span>{status}</span>
          </div>
        )}

        {/* Download */}
        {downloadUrl && (
          <div className="mb-6">
            <button
              onClick={handleDownload}
              disabled={downloading}
              className="flex items-center justify-center gap-2 w-full bg-green-600 dark:bg-green-500 text-white py-3 px-6 rounded-lg font-medium hover:bg-green-700 dark:hover:bg-green-600 disabled:bg-green-400 dark:disabled:bg-green-400 disabled:cursor-not-allowed transition-colors"
            >
              {downloading ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  Downloading...
                </>
              ) : (
                <>
                  <Download className="w-5 h-5" /> Download Excel
                </>
              )}
            </button>
          </div>
        )}



        {/* Log */}
        {log.length > 0 && (
          <div className="bg-gray-50 dark:bg-gray-700 rounded-lg p-4 border border-gray-200 dark:border-gray-600 max-h-64 overflow-y-auto">
            <h3 className="text-sm font-semibold text-gray-700 dark:text-gray-200 mb-2">
              Processing Log:
            </h3>
            <div className="space-y-1">
              {log.map((entry, idx) => (
                <div
                  key={idx}
                  className="text-xs font-mono text-gray-600 dark:text-gray-300"
                >
                  {entry}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="mt-4 pt-6 border-t border-gray-200 dark:border-gray-700 text-center">
          <p className="text-gray-600 dark:text-gray-300 text-sm flex items-center justify-center gap-2">
            Made with <Heart className="w-4 h-4 text-red-500 fill-red-500" /> by
            Eric
          </p>
          <p className="text-gray-500 dark:text-gray-400 text-xs mt-2">
            © {new Date().getFullYear()} Traffic Count Modifier. All rights
            reserved.
          </p>
        </div>
      </div>
    </div>
  );
}
