import React, { useState, useRef } from "react";
import { Upload, FileText, Image as ImageIcon, Download, Loader2, CheckCircle2, AlertCircle, Plus, ExternalLink } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

export default function App() {
  const [template, setTemplate] = useState<File | null>(null);
  const [images, setImages] = useState<File[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const templateInputRef = useRef<HTMLInputElement>(null);
  const imagesInputRef = useRef<HTMLInputElement>(null);

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setTemplate(e.target.files[0]);
      setError(null);
    }
  };

  const handleImagesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setImages(prev => [...prev, ...Array.from(e.target.files!)]);
      setError(null);
    }
  };

  const generateWord = async () => {
    if (!template || images.length === 0) {
      setError("請提供範例 Word 檔與至少一張圖片");
      return;
    }

    setIsGenerating(true);
    setError(null);
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("template", template);
    images.forEach(img => formData.append("images", img));

    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errorMessage = `伺服器錯誤 (${response.status})`;
        try {
          const errData = await response.json();
          errorMessage = errData.error || errorMessage;
        } catch (e) {
          // If response is not JSON (e.g. HTML error page)
          const text = await response.text();
          console.error("Non-JSON error response:", text);
          if (text.includes("<!doctype")) {
            errorMessage = "伺服器返回了 HTML 頁面而非 JSON。這通常意味著路徑錯誤或伺服器崩潰。";
          }
        }
        throw new Error(errorMessage);
      }

      const responseClone = response.clone();
      const data = await response.json().catch(async (e) => {
        const text = await responseClone.text();
        console.error("Failed to parse JSON response:", text);
        if (text.includes("<!doctype html>") || text.includes("<html")) {
          throw new Error("伺服器返回了 HTML 格式的錯誤。這通常表示路徑錯誤、伺服器重啟或處理超時。");
        }
        throw new Error("無法解析伺服器回應。可能是伺服器發生了未預期的錯誤。");
      });
      setDownloadUrl(data.downloadUrl);
    } catch (err: any) {
      console.error("Generate error:", err);
      setError(err.message);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleTemplateDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith(".docx")) {
        setTemplate(file);
        setError(null);
      } else {
        setError("請上傳 .docx 格式的 Word 檔案");
      }
    }
  };

  const handleImagesDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files) {
      const newFiles = Array.from(e.dataTransfer.files as FileList).filter(f => (f as File).type.startsWith("image/"));
      if (newFiles.length > 0) {
        setImages(prev => [...prev, ...newFiles]);
        setError(null);
      }
    }
  };

  const reset = () => {
    setTemplate(null);
    setImages([]);
    setDownloadUrl(null);
    setError(null);
    if (templateInputRef.current) templateInputRef.current.value = "";
    if (imagesInputRef.current) imagesInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-neutral-50 text-neutral-900 font-sans p-6 md:p-12">
      <div className="max-w-4xl mx-auto">
        <header className="mb-12 flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div>
            <h1 className="text-4xl font-bold tracking-tight mb-2">Word 表格圖片填充器</h1>
            <p className="text-neutral-500">上傳含表格的 Word 範本，自動將圖片依序填入並無限複製表格。</p>
          </div>
          <button
            onClick={() => window.open(window.location.href, '_blank')}
            className="flex items-center gap-2 text-sm font-medium text-neutral-500 hover:text-neutral-900 transition-colors bg-white px-4 py-2 rounded-full border border-neutral-200 shadow-sm"
          >
            <ExternalLink className="w-4 h-4" />
            以新分頁開啟
          </button>
        </header>

        <div className="grid md:grid-cols-2 gap-8">
          {/* Step 1: Template */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-neutral-200">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 font-bold">1</div>
              <h2 className="text-lg font-semibold">上傳範例 Word</h2>
            </div>
            
            <div 
              onClick={() => templateInputRef.current?.click()}
              onDragOver={handleDragOver}
              onDrop={handleTemplateDrop}
              className={`border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center cursor-pointer transition-colors ${
                template ? "border-blue-400 bg-blue-50" : "border-neutral-200 hover:border-blue-300 hover:bg-neutral-50"
              }`}
            >
              <input 
                type="file" 
                ref={templateInputRef} 
                onChange={handleTemplateUpload} 
                accept=".docx" 
                className="hidden" 
              />
              {template ? (
                <>
                  <FileText className="w-12 h-12 text-blue-500 mb-2" />
                  <span className="text-sm font-medium text-blue-700">{template.name}</span>
                </>
              ) : (
                <>
                  <Upload className="w-12 h-12 text-neutral-300 mb-2" />
                  <span className="text-sm text-neutral-500 italic">點擊或拖放 .docx 檔案</span>
                </>
              )}
            </div>
            
            <p className="mt-4 text-xs text-neutral-400 leading-relaxed italic">
              提示：範本中的表格內應包含 <code className="bg-neutral-100 px-1 rounded italic font-bold">[圖片]</code> 文字標記。
            </p>
          </section>

          {/* Step 2: Images */}
          <section className="bg-white p-6 rounded-2xl shadow-sm border border-neutral-200">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 rounded-full bg-green-100 flex items-center justify-center text-green-600 font-bold">2</div>
              <h2 className="text-lg font-semibold">選取/上傳圖片</h2>
            </div>

            <div 
              onClick={() => imagesInputRef.current?.click()}
              onDragOver={handleDragOver}
              onDrop={handleImagesDrop}
              className={`border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center cursor-pointer transition-colors ${
                images.length > 0 ? "border-green-400 bg-green-50" : "border-neutral-200 hover:border-green-300 hover:bg-neutral-50"
              }`}
            >
              <input 
                type="file" 
                ref={imagesInputRef} 
                onChange={handleImagesUpload} 
                multiple 
                accept="image/*" 
                className="hidden" 
              />
              {images.length > 0 ? (
                <>
                  <ImageIcon className="w-12 h-12 text-green-500 mb-2" />
                  <span className="text-sm font-medium text-green-700 italic">已選取 {images.length} 張圖片</span>
                </>
              ) : (
                <>
                  <Plus className="w-12 h-12 text-neutral-300 mb-2" />
                  <span className="text-sm text-neutral-500 italic">選取多張圖片</span>
                </>
              )}
            </div>

            {images.length > 0 && (
              <div className="mt-4 flex flex-wrap gap-2 max-h-32 overflow-y-auto">
                {images.slice(0, 10).map((img, i) => (
                  <div key={i} className="w-10 h-10 rounded bg-neutral-100 overflow-hidden flex items-center justify-center text-xs text-neutral-400">
                    <img src={URL.createObjectURL(img)} className="w-full h-full object-cover" />
                  </div>
                ))}
                {images.length > 10 && <div className="w-10 h-10 flex items-center justify-center text-xs text-neutral-400">+{images.length - 10}</div>}
              </div>
            )}
          </section>
        </div>

        {/* Actions */}
        <div className="mt-12 flex flex-col items-center">
          <AnimatePresence mode="wait">
            {!downloadUrl ? (
              <motion.button
                key="generate"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                disabled={isGenerating || !template || images.length === 0}
                onClick={generateWord}
                className="px-12 py-4 bg-neutral-900 text-white rounded-full font-bold shadow-lg hover:bg-neutral-800 disabled:bg-neutral-300 disabled:cursor-not-allowed transition-all flex items-center gap-2 group"
              >
                {isGenerating ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    正在處理範本與圖片...
                  </>
                ) : (
                  <>
                    <FileText className="w-5 h-5 group-hover:scale-110 transition-transform" />
                    生成 Word 檔案
                  </>
                )}
              </motion.button>
            ) : (
              <div className="flex flex-col items-center gap-4">
                <motion.div
                  key="success"
                  initial={{ opacity: 0, scale: 0.9 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="flex items-center gap-2 text-green-600 font-medium"
                >
                  <CheckCircle2 className="w-6 h-6" />
                  成功生成！
                </motion.div>
                <div className="flex gap-4">
                  <a
                    href={downloadUrl}
                    className="px-12 py-4 bg-green-600 text-white rounded-full font-bold shadow-lg hover:bg-green-700 flex items-center gap-2"
                  >
                    <Download className="w-5 h-5" />
                    立即下載
                  </a>
                  <button
                    onClick={reset}
                    className="px-8 py-4 bg-white border border-neutral-200 text-neutral-600 rounded-full font-bold hover:bg-neutral-50 transition-colors"
                  >
                    重新開始
                  </button>
                </div>
              </div>
            )}
          </AnimatePresence>

          {error && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="mt-6 flex items-center gap-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-100"
            >
              <AlertCircle className="w-4 h-4" />
              <span className="text-sm font-medium">{error}</span>
            </motion.div>
          )}
        </div>

        {/* Instructions */}
        <div className="mt-24 border-t border-neutral-200 pt-12 pb-24">
          <h3 className="text-sm font-bold uppercase tracking-widest text-neutral-400 mb-6 text-center italic">使用說明與提示</h3>
          <div className="grid md:grid-cols-3 gap-8 text-sm text-neutral-600">
            <div>
              <h4 className="font-bold text-neutral-900 mb-2 italic">1. 設置標記</h4>
              <p className="italic">在 Word 的表格格子內輸入 <span className="font-bold text-neutral-900">[圖片]</span>。每個格子放一個標記，程式會自動識別其座標位置。</p>
            </div>
            <div>
              <h4 className="font-bold text-neutral-900 mb-2 italic">2. 自動表格複製</h4>
              <p className="italic">如果你的範本只有一個表格（內含兩個 [圖片] 標記），但你上傳了 10 張照片，程式會自動複製 5 個相同的表格並填滿所有圖片。</p>
            </div>
            <div>
              <h4 className="font-bold text-neutral-900 mb-2 italic">3. 保持格式</h4>
              <p className="italic">填充後的圖片會儘量維持標記處的大小。系統會自動刪除原本的 [圖片] 文字，只留下清晰的照片。</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
