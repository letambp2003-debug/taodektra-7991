import { useState, useRef, ChangeEvent, createContext, useContext } from "react";
import { GoogleGenAI } from "@google/genai";
import { asBlob } from "html-docx-js-typescript";
import { Upload, FileText, Loader2, CheckCircle, AlertCircle, RefreshCw, Copy, Check, Edit2, Save, FileSpreadsheet, FileDown, Key, ExternalLink, X } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import rehypeRaw from "rehype-raw";
import rehypeKatex from "rehype-katex";
import "katex/dist/katex.min.css";
import * as XLSX from "xlsx";
import * as mammoth from "mammoth";

// Danh sách mô hình theo thứ tự ưu tiên từ thông minh/trọng số cao nhất xuống đến các model nhẹ và phản hồi nhanh nhất
const MODEL_PRIORITY_LIST: string[] = [
  "gemini-3.8-flash",
  "gemini-3.7-flash",
  "gemini-3.6-flash",
  "gemini-3.1-pro",
  "gemini-3.5-flash",
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-2.5-pro",
  "gemini-2.5-flash",
  "gemini-2.5-flash-lite",
  "gemini-3-flash",
  "gemini-2-flash",
  "gemini-2-flash-lite",
];

// Ngưỡng ký tự để kích hoạt Context Caching Engine (~32k tokens theo quy chuẩn của Google Cloud)
const CONTEXT_CACHE_CHAR_THRESHOLD = 110000;

// Hàm giải quyết danh sách tên định danh tương thích (bao gồm phiên bản preview nếu có)
const getModelCandidates = (modelName: string): string[] => {
  const candidates = [modelName];
  if (modelName === "gemini-3.1-pro") {
    candidates.push("gemini-3.1-pro-preview");
  } else if (modelName === "gemini-3-flash") {
    candidates.push("gemini-3-flash-preview");
  } else if (modelName === "gemini-2-flash") {
    candidates.push("gemini-2.0-flash", "gemini-2.0-flash-001");
  } else if (modelName === "gemini-2-flash-lite") {
    candidates.push("gemini-2.0-flash-lite", "gemini-2.0-flash-lite-preview-02-05");
  }
  return candidates;
};

// Context to track if we are inside an SVG element
const SvgContext = createContext(false);

const processHtmlForDocx = async (element: HTMLElement): Promise<string> => {
  const clone = element.cloneNode(true) as HTMLElement;
  
  // Replace KaTeX elements with raw LaTeX for Word conversion
  const katexElements = clone.querySelectorAll('.katex');
  katexElements.forEach(el => {
    const annotation = el.querySelector('annotation[encoding="application/x-tex"]');
    if (annotation) {
      const texText = annotation.textContent;
      const textNode = document.createTextNode('$' + texText + '$');
      el.parentNode?.replaceChild(textNode, el);
    }
  });

  // Convert external images to base64
  const images = clone.querySelectorAll('img');
  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    if (img.src && !img.src.startsWith('data:')) {
      try {
        const response = await fetch(img.src);
        const blob = await response.blob();
        const base64 = await new Promise<string>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(reader.result as string);
          reader.readAsDataURL(blob);
        });
        img.src = base64;
      } catch (e) {
        console.error('Failed to fetch image for export', e);
      }
    }
  }

  // Convert SVG to base64 PNG
  const svgs = clone.querySelectorAll('svg');
  for (let i = 0; i < svgs.length; i++) {
    const svg = svgs[i];
    try {
      const svgData = new XMLSerializer().serializeToString(svg);
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const img = new Image();
      
      const rect = svg.getBoundingClientRect();
      const width = parseFloat(svg.getAttribute('width') || '0') || rect.width || 300;
      const height = parseFloat(svg.getAttribute('height') || '0') || rect.height || 150;
      canvas.width = width;
      canvas.height = height;

      const base64Svg = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
      
      await new Promise<void>((resolve, reject) => {
        img.onload = () => {
          if (ctx) {
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, width, height);
            ctx.drawImage(img, 0, 0);
          }
          resolve();
        };
        img.onerror = reject;
        img.src = base64Svg;
      });

      const pngBase64 = canvas.toDataURL('image/png');
      const newImg = document.createElement('img');
      newImg.src = pngBase64;
      newImg.width = width;
      newImg.height = height;
      svg.parentNode?.replaceChild(newImg, svg);
    } catch (e) {
      console.error('Failed to convert SVG to PNG', e);
    }
  }

  return clone.innerHTML;
};

export default function App() {
  const [fileName, setFileName] = useState("");
  const [fileData, setFileData] = useState<string | null>(null);
  const [fileText, setFileText] = useState<string | null>(null);
  const [specFileName, setSpecFileName] = useState("");
  const [specFileData, setSpecFileData] = useState<string | null>(null);
  const [specFileText, setSpecFileText] = useState<string | null>(null);
  const [manualInput, setManualInput] = useState("");
  const [mimeType, setMimeType] = useState<string>("");
  const [specMimeType, setSpecMimeType] = useState<string>("");
  const [examType, setExamType] = useState("Giữa kì 2");
  const [duration, setDuration] = useState(90);
  const [counts, setCounts] = useState({
    multipleChoice: 12,
    trueFalse: 2,
    shortAnswer: 4,
    essay: 3
  });
  const [points, setPoints] = useState({
    multipleChoice: 3.0,
    trueFalse: 2.0,
    shortAnswer: 2.0,
    essay: 3.0
  });
  const [ratios, setRatios] = useState({
    know: 40,
    understand: 30,
    apply: 30
  });
  const [lessonPlan, setLessonPlan] = useState("");
  const [matrix, setMatrix] = useState("");
  const [specTable, setSpecTable] = useState("");
  const [exam, setExam] = useState("");
  const [exam2, setExam2] = useState("");
  const [loading, setLoading] = useState(false);
  const [loadingStatus, setLoadingStatus] = useState("Đang xử lý...");
  const [activeModel, setActiveModel] = useState<string | null>(null);
  const [isCached, setIsCached] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [apiKey, setApiKey] = useState<string>(() => localStorage.getItem("GEMINI_API_KEY") || (process.env.GEMINI_API_KEY as string) || "");
  const [showApiKeyModal, setShowApiKeyModal] = useState(false);
  const [tempApiKey, setTempApiKey] = useState<string>(() => localStorage.getItem("GEMINI_API_KEY") || (process.env.GEMINI_API_KEY as string) || "");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const specFileInputRef = useRef<HTMLInputElement>(null);
  const resultsRef = useRef<HTMLDivElement>(null);

  const handleSaveApiKey = () => {
    const trimmed = tempApiKey.trim();
    setApiKey(trimmed);
    if (trimmed) {
      localStorage.setItem("GEMINI_API_KEY", trimmed);
    } else {
      localStorage.removeItem("GEMINI_API_KEY");
    }
    setShowApiKeyModal(false);
  };

  const examTypes = [
    "Giữa kì 1",
    "Cuối kì 1",
    "Giữa kì 2",
    "Cuối kì 2",
    "Kiểm tra thường xuyên",
    "Khác"
  ];

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 10 * 1024 * 1024) { // 10MB limit
        setError("File quá lớn. Vui lòng tải lên file nhỏ hơn 10MB.");
        return;
      }
      
      setFileName(file.name);
      setMimeType(file.type);
      setFileText(null);
      setError(null);

      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64String = reader.result as string;
        const base64Data = base64String.split(',')[1];
        setFileData(base64Data);
        
        // Handle .docx files
        if (file.name.endsWith('.docx')) {
          try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            setFileText(result.value);
          } catch (err) {
            console.error("Error extracting text from docx:", err);
            setError("Không thể trích xuất văn bản từ file .docx. Vui lòng thử lại hoặc chuyển sang PDF.");
          }
        }
        
        // If it's a text file, also set manual input
        if (file.type.startsWith('text/') || file.name.endsWith('.txt') || file.name.endsWith('.md')) {
          const textReader = new FileReader();
          textReader.onload = (e) => {
            setManualInput(e.target?.result as string);
          };
          textReader.readAsText(file);
        }
      };
      reader.onerror = () => {
        setError("Lỗi khi đọc file. Vui lòng thử lại.");
      };
      reader.readAsDataURL(file);
    }
  };

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  const handleSpecFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 10 * 1024 * 1024) {
        setError("File quá lớn. Vui lòng tải lên file nhỏ hơn 10MB.");
        return;
      }
      
      setSpecFileName(file.name);
      setSpecMimeType(file.type);
      setSpecFileText(null);
      setError(null);

      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64String = reader.result as string;
        const base64Data = base64String.split(',')[1];
        setSpecFileData(base64Data);

        // Handle .docx files
        if (file.name.endsWith('.docx')) {
          try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            setSpecFileText(result.value);
          } catch (err) {
            console.error("Error extracting text from docx:", err);
            setError("Không thể trích xuất văn bản từ file .docx đặc tả.");
          }
        }
      };
      reader.onerror = () => {
        setError("Lỗi khi đọc file đặc tả. Vui lòng thử lại.");
      };
      reader.readAsDataURL(file);
    }
  };

  const triggerSpecFileInput = () => {
    specFileInputRef.current?.click();
  };

  const handleExportWordAll = async () => {
    if (!resultsRef.current) return;
    
    const mathTypeNote = `<p style="color: red; font-weight: bold; margin-bottom: 20px; font-style: italic;">
      Hãy sử dụng MathType để chuyển các công thức Toán học, Hóa học. Cách làm cụ thể:<br/>
      Bôi đen cả dòng chứa công thức (thường có dấu $), bấm MathType, chọn Toogle Text
    </p>`;
    
    let htmlContent = mathTypeNote;
    const cards = resultsRef.current.querySelectorAll('.prose');
    const titles = resultsRef.current.querySelectorAll('h2');
    
    for (let index = 0; index < cards.length; index++) {
      const card = cards[index] as HTMLElement;
      const title = titles[index]?.textContent || "";
      const processedHtml = await processHtmlForDocx(card);
      
      htmlContent += `<h2 style="text-align: center; color: #1e3a8a; font-size: 16pt; text-transform: uppercase; margin-bottom: 20px;">${title}</h2>`;
      htmlContent += processedHtml;
      if (index < cards.length - 1) {
        htmlContent += `<br clear="all" style="page-break-before:always" />`;
      }
    }

    const header = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Hồ sơ dạy học</title>
    <style>
      body { 
        font-family: 'Times New Roman', Times, serif; 
        font-size: 12pt; 
        mso-ascii-font-family: 'Times New Roman'; 
        mso-hansi-font-family: 'Times New Roman'; 
        mso-bidi-font-family: 'Times New Roman';
      }
      table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
      th, td { border: 1px solid black; padding: 8px; text-align: left; font-family: 'Times New Roman', Times, serif; }
      th { background-color: #f2f2f2; font-weight: bold; text-align: center; }
      h1, h2, h3, h4, h5, h6 { 
        font-family: 'Times New Roman', Times, serif; 
        mso-ascii-font-family: 'Times New Roman'; 
        mso-hansi-font-family: 'Times New Roman'; 
      }
      p { line-height: 1.5; margin-bottom: 10px; font-family: 'Times New Roman', Times, serif; }
      img { max-width: 100%; height: auto; }
    </style>
    </head><body>`;
    const footer = "</body></html>";
    const sourceHTML = header + htmlContent + footer;
    
    try {
      const blob = await asBlob(sourceHTML) as Blob;
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `Ho_So_Day_Hoc_${examType.replace(/\s+/g, '_')}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Export to DOCX failed:", error);
      alert("Đã xảy ra lỗi khi xuất file Word. Vui lòng thử lại.");
    }
  };

  const generateAll = async () => {
    const validationErrors = [];
    
    if (!fileData && !manualInput) {
      validationErrors.push("• Vui lòng tải lên file hoặc nhập nội dung Phụ lục 3.");
    }
    if (duration <= 0) {
      validationErrors.push("• Thời gian làm bài phải lớn hơn 0 phút.");
    }
    if (ratios.know + ratios.understand + ratios.apply !== 100) {
      validationErrors.push(`• Tổng tỉ lệ điểm phải bằng 100% (hiện tại là ${ratios.know + ratios.understand + ratios.apply}%).`);
    }
    if (counts.multipleChoice + counts.trueFalse + counts.shortAnswer + counts.essay <= 0) {
      validationErrors.push("• Tổng số câu hỏi phải lớn hơn 0.");
    }
    if (points.multipleChoice + points.trueFalse + points.shortAnswer + points.essay <= 0) {
      validationErrors.push("• Tổng số điểm phải lớn hơn 0.");
    }

    const effectiveApiKey = apiKey.trim() || localStorage.getItem("GEMINI_API_KEY") || (process.env.GEMINI_API_KEY as string) || "";
    if (!effectiveApiKey) {
      validationErrors.push("• Vui lòng nhập Gemini API Key để tiếp tục. Nhấn vào nút 'Cài đặt API Key' ở đầu trang.");
      setShowApiKeyModal(true);
    }

    if (validationErrors.length > 0) {
      setError(validationErrors.join("\n"));
      return;
    }

    setLoading(true);
    setLoadingStatus("Đang kiểm tra và chuẩn bị dữ liệu...");
    setActiveModel(null);
    setIsCached(false);
    setError(null);
    setLessonPlan("");
    setMatrix("");
    setSpecTable("");
    setExam("");
    setExam2("");

    const totalPoints = points.multipleChoice + points.trueFalse + points.shortAnswer + points.essay;
    const tnkqPoints = points.multipleChoice + points.trueFalse + points.shortAnswer;
    const essayPoints = points.essay;

    try {
      const prompt = `
Bạn là chuyên gia trợ lý tạo đề kiểm tra theo chuẩn Công văn 7991/BGDĐT-GDTrH của Bộ GDĐT, phục vụ công tác giảng dạy của GVBM: Lê Tâm - Trường THCS Quang Trung.
Loại đề cần tạo: ${examType}.
Thời gian làm bài: ${duration} phút.

Dựa vào nội dung Phụ lục 3 và Tài liệu mức độ đặc tả mà giáo viên cung cấp, hãy tạo đầy đủ 5 phần sau cho đề ${examType} (thời gian ${duration} phút) theo đúng cấu trúc chuẩn của Bộ GDĐT như sau:

PHÂN TÍCH PHỤ LỤC 3 (Yêu cầu ngầm định):
- BẮT BUỘC phải phân tích kỹ Phụ lục 3 để xác định số tiết học của từng nội dung/đơn vị kiến thức.
- Nội dung nào có số tiết học nhiều hơn thì BẮT BUỘC phải được phân bổ nhiều câu hỏi và nhiều điểm hơn trong Ma trận (Phần 2) và Bảng đặc tả (Phần 3).

PHẦN 1: KẾ HOẠCH BÀI DẠY
Dựa vào nội dung Phụ lục 3, hãy tạo khung kế hoạch bài dạy với cấu trúc chính xác như sau:
I. MỤC TIÊU
1. Về kiến thức
2. Về năng lực
a. Năng lực chung
b. Năng lực đặc thù
3. Phẩm chất
II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
1. Thiết bị dạy học
2. Học liệu
III. TIẾN TRÌNH DẠY HỌC

PHẦN 2: MA TRẬN ĐỀ
Yêu cầu tạo bảng bằng HTML (để hỗ trợ rowspan và colspan) với cấu trúc header chính xác như sau:
<table border="1" style="border-collapse: collapse; width: 100%; text-align: center;">
  <thead>
    <tr>
      <th rowspan="3">TT</th>
      <th rowspan="3">Chủ đề/ Chương</th>
      <th rowspan="3">Nội dung/ đơn vị kiến thức</th>
      <th colspan="12">Mức độ đánh giá</th>
      <th colspan="3" rowspan="2">Tổng</th>
      <th rowspan="3">Tỉ lệ % điểm</th>
    </tr>
    <tr>
      <th colspan="9">TNKQ</th>
      <th colspan="3">Tự luận</th>
    </tr>
    <tr>
      <th colspan="3">Nhiều lựa chọn</th>
      <th colspan="3">"Đúng – Sai"</th>
      <th colspan="3">Trả lời ngắn</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
    </tr>
    <tr>
      <th></th><th></th><th></th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th></th><th></th><th></th>
      <th></th><th></th><th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <!-- AI: Tự động điền các dòng dữ liệu (tr) và ô (td) tương ứng vào đây -->
  </tbody>
  <tfoot>
    <!-- AI: BẮT BUỘC phải có 3 dòng cuối cùng để tổng kết:
      1. Dòng "Tổng số câu": Tính tổng số câu cho từng cột mức độ đánh giá (Biết, Hiểu, Vận dụng) của từng phần.
      2. Dòng "Tổng số điểm": Tính tổng số điểm tương ứng cho từng cột. Mọi cột nếu không có câu hỏi thì để là 0.
      3. Dòng "Tỉ lệ %": Tính tỉ lệ phần trăm điểm cho từng cột (Biết, Hiểu, Vận dụng) dựa trên tổng ${totalPoints} điểm. Phần trăm này BẮT BUỘC phải KHỚP CHÍNH XÁC với ${ratios.know}-${ratios.understand}-${ratios.apply}. Rất quan trọng: nếu vòng Đúng/Sai và Trả lời ngắn bằng 0 câu, bạn BẮT BUỘC phải dồn Toàn Bộ % Biết (${ratios.know}%) và % Hiểu (${ratios.understand}%) vào cột Tỉ lệ % tương ứng của phần "Nhiều lựa chọn". Tổng tỉ lệ cuối cùng phải luôn là 100%.
    -->
  </tfoot>
</table>

PHẦN 3: BẢNG ĐẶC TẢ
Yêu cầu tạo bảng bằng HTML với cấu trúc header chính xác như sau:
<table border="1" style="border-collapse: collapse; width: 100%; text-align: center;">
  <thead>
    <tr>
      <th rowspan="3">TT</th>
      <th rowspan="3">Chủ đề/ Chương</th>
      <th rowspan="3">Nội dung/ đơn vị kiến thức</th>
      <th rowspan="3">Yêu cầu cần đạt</th>
      <th colspan="12">Số câu hỏi ở các mức độ đánh giá</th>
    </tr>
    <tr>
      <th colspan="9">TNKQ</th>
      <th colspan="3">Tự luận</th>
    </tr>
    <tr>
      <th colspan="3">Nhiều lựa chọn</th>
      <th colspan="3">"Đúng – Sai"</th>
      <th colspan="3">Trả lời ngắn</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
    </tr>
    <tr>
      <th></th><th></th><th></th><th></th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th>Biết</th><th>Hiểu</th><th>Vận dụng</th>
      <th></th><th></th><th></th>
    </tr>
  </thead>
  <tbody>
    <!-- AI: Tự động điền các dòng dữ liệu (tr) và ô (td) tương ứng vào đây -->
  </tbody>
  <tfoot>
    <!-- AI: BẮT BUỘC phải có 3 dòng cuối cùng để tổng kết (tương tự Ma trận):
      1. Dòng "Tổng số câu": Tính tổng số câu cho từng cột mức độ đánh giá.
      2. Dòng "Tổng số điểm": Tính tổng số điểm tương ứng cho từng cột. Mọi cột nếu không có câu hỏi thì để là 0.
      3. Dòng "Tỉ lệ %": Tính tỉ lệ phần trăm điểm cho từng cột dựa trên tổng ${totalPoints} điểm. Phần trăm này BẮT BUỘC phải KHỚP CHÍNH XÁC với ${ratios.know}-${ratios.understand}-${ratios.apply}. Rất quan trọng: nếu vòng Đúng/Sai và Trả lời ngắn bằng 0 câu, bạn BẮT BUỘC phải dồn Toàn Bộ % Biết (${ratios.know}%) và % Hiểu (${ratios.understand}%) vào cột Tỉ lệ % tương ứng của phần "Nhiều lựa chọn" sao cho tổng tỉ lệ cuối cùng phải luôn là 100%.
    -->
  </tfoot>
</table>

PHẦN 4: ĐỀ KIỂM TRA SỐ 1
Cấu trúc đề gồm các phần:
${[
  counts.multipleChoice > 0 ? `- PHẦN I: Trắc nghiệm nhiều phương án lựa chọn (BẮT BUỘC tạo ĐÚNG ${counts.multipleChoice} câu, tổng ${points.multipleChoice} điểm)` : "",
  counts.trueFalse > 0 ? `- PHẦN II: Trắc nghiệm Đúng - Sai (BẮT BUỘC tạo ĐÚNG ${counts.trueFalse} câu, tổng ${points.trueFalse} điểm)` : "",
  counts.shortAnswer > 0 ? `- PHẦN III: Trắc nghiệm Trả lời ngắn (BẮT BUỘC tạo ĐÚNG ${counts.shortAnswer} câu, tổng ${points.shortAnswer} điểm)` : "",
  counts.essay > 0 ? `- PHẦN IV: Tự luận (BẮT BUỘC tạo ĐÚNG ${counts.essay} câu, tổng ${points.essay} điểm)` : ""
].filter(Boolean).join("\n")}
- ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT

PHẦN 5: ĐỀ KIỂM TRA SỐ 2
- Yêu cầu: Tạo một đề kiểm tra thứ 2 tương đương hoàn toàn với Đề số 1 về cấu trúc, độ khó, và bám sát chính xác Ma trận (Phần 2) và Bảng đặc tả (Phần 3).
- Cấu trúc: Giống hệt Đề số 1.
- Nội dung: Các câu hỏi phải được thay đổi số liệu, ngữ cảnh, hoặc cách hỏi nhưng vẫn giữ nguyên dạng bài và mức độ nhận thức tương ứng với Đề 1.
- ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT cho Đề số 2.

QUY ĐỊNH BẮT BUỘC CHO CẢ 2 ĐỀ:
1. Tổng số câu hỏi trong đề phải bằng ĐÚNG ${counts.multipleChoice + counts.trueFalse + counts.shortAnswer + counts.essay} câu. Tổng điểm là ${totalPoints} điểm. (Nếu quy định 0 câu ở phần nào, tuyệt đối không tạo phần đó trong đề, và điền 0 hoặc bỏ trống các nội dung trong ma trận/đặc tả ở phần đó).
2. Tỉ lệ điểm theo mức độ nhận thức (Biết - Hiểu - Vận dụng) phải ĐÚNG theo tỉ lệ do người dùng quy định: ${ratios.know}% - ${ratios.understand}% - ${ratios.apply}% (Tương ứng hệ số điểm chuẩn xác: ${(ratios.know * totalPoints / 100).toFixed(2)}đ - ${(ratios.understand * totalPoints / 100).toFixed(2)}đ - ${(ratios.apply * totalPoints / 100).toFixed(2)}đ). Chỉ số phần trăm này phải được hiển thị SIÊU RÕ RÀNG ở dòng tổng kết cuối của Bảng Ma Trận và Bảng Đặc Tả, luôn giữ nguyên cấu trúc dù nhập thế nào.
3. Phân bố điểm chính xác giữa các phần: Trắc nghiệm KQ BẮT BUỘC là ${tnkqPoints} điểm và Tự luận BẮT BUỘC là ${essayPoints} điểm. Cụ thể: Nhiều lựa chọn (${points.multipleChoice}đ), Đúng-Sai (${points.trueFalse}đ), Trả lời ngắn (${points.shortAnswer}đ), Tự luận (${points.essay}đ). ĐẶC BIỆT CHÚ Ý VÀO MA TRẬN VÀ BẢNG ĐẶC TẢ: Nếu số lượng câu hỏi phần "Đúng-Sai" và "Trả lời ngắn" bằng 0, bạn BẮT BUỘC phải dồn toàn bộ điểm và tỉ lệ % ở các mức độ tương ứng của phần TNKQ (ví dụ toàn bộ ${ratios.know}% Biết, ${ratios.understand}% Hiểu hoặc bất kỳ tỉ lệ nào đang tính cho TNKQ) điền hết sạch vào cột "Nhiều lựa chọn". Mọi thông số của 2 phần có số câu bằng 0 sẽ bị vô hiệu hóa (xóa hoặc để trống hoàn toàn). Phân bố cuối phải khép kín tuyệt đối tỉ lệ tổng.
4. TRÌNH BÀY ĐỀ: Trong phần trắc nghiệm nhiều phương án lựa chọn, BẮT BUỘC mỗi đáp án A, B, C, D phải nằm trên 1 dòng riêng biệt (tổng cộng 4 dòng cho 4 đáp án). Tuyệt đối không viết liền nhau trên cùng một dòng. (Lưu ý: Thêm 2 dấu cách ở cuối mỗi dòng hoặc dùng 1 dòng trống giữa các đáp án để Markdown hiển thị xuống dòng chính xác). Các câu hỏi cũng phải được tách biệt rõ ràng bằng dòng trống.
5. ĐỐI VỚI PHẦN TRẮC NGHIỆM TRẢ LỜI NGẮN: Đáp án BẮT BUỘC phải cực kỳ ngắn gọn, CHỈ BAO GỒM số, đơn vị, hoặc một công thức thật ngắn. Tuyệt đối không giải thích dài dòng hay viết thành câu hoàn chỉnh trong phần đáp án của học sinh.
6. Bảng đặc tả (PHẦN 3) - CỘT "YÊU CẦU CẦN ĐẠT": BẮT BUỘC phải phân tích và ghi thật chi tiết. PHẢI TÁCH RIÊNG RÕ RÀNG từng mức độ đánh giá có trong câu hỏi/bài học đó (Ví dụ: cống dòng ghi "- Nhận biết: [mô tả chi tiết...]"; xuống dòng ghi "- Thông hiểu: [mô tả chi tiết...]"; xuống dòng ghi "- Vận dụng: [mô tả chi tiết...]"). TUYỆT ĐỐI KHÔNG ghi chung chung, không được gộp chung các mức độ vào nhau mà phải tách bạch cụ thể từng dòng tương ứng với mức độ đó.
7. TRONG BẢNG ĐẶC TẢ, phải ghi rõ số thứ tự câu hỏi tương ứng vào các ô mức độ (Ví dụ: "Câu 1", "Câu 19a", "Câu 20",...).
8. Ma trận đề (PHẦN 2) và Bảng đặc tả (PHẦN 3) phải thống nhất hoàn toàn với số lượng câu hỏi và phân bổ điểm thực tế trong đề (PHẦN 4).
9. KIỂM TRA TOÁN HỌC NGHIÊM NGẶT - BẢO ĐẢM TỈ LỆ AN TOÀN: Khi cộng tổng điểm của tất cả các dòng, nội dung tương ứng của mỗi cột (Biết, Hiểu, Vận dụng) trong Ma trận (PHẦN 2) và Bảng đặc tả (PHẦN 3), TỔNG CUỐI CÙNG BẮT BUỘC PHẢI KHỚP TUYỆT ĐỐI VỚI MỨC NGƯỜI DÙNG QUY ĐỊNH: Biết=${ratios.know}%, Hiểu=${ratios.understand}%, Vận dụng=${ratios.apply}% (tức là ${(ratios.know * totalPoints / 100).toFixed(2)} điểm, ${(ratios.understand * totalPoints / 100).toFixed(2)} điểm, ${(ratios.apply * totalPoints / 100).toFixed(2)} điểm). Cho dù điểm bị lẻ thế nào, AI cũng BẮT BUỘC phải điều tiết hệ số mỗi câu sao cho khi cộng lại không được sai lệch dù chỉ là 0.01 điểm. Bảng luôn phải chia cột chính xác kể cả khi không có tự luận.
10. HƯỚNG DẪN CHẤM:
    - Đối với TNKQ (Phần I, II, III): Trình bày bảng đáp án gồm 2 dòng (Dòng 1: Câu; Dòng 2: Đáp án).
    - Đối với Tự luận (Phần IV): Trình bày bảng gồm 3 cột (Cột 1: Câu; Cột 2: Nội dung đáp án/Hướng dẫn giải chi tiết; Cột 3: Điểm).
    - Phải trình bày chi tiết từng bước giải và cho điểm cụ thể cho từng bước đó. Riêng đối với các bài toán hình học (nếu có), quy định vẽ hình đúng được tính 0,25 điểm.
11. TỔNG ĐIỂM TOÀN ĐỀ: Phải luôn là ${totalPoints} điểm (TNKQ ${tnkqPoints}đ + Tự luận ${essayPoints}đ).
12. ĐỘ KHÓ VÀ THỜI GIAN: Đảm bảo độ khó và khối lượng kiến thức của đề thi phù hợp tuyệt đối với thời gian làm bài là ${duration} phút. Các câu hỏi phải được phân bổ sao cho học sinh trung bình khá có thể hoàn thành đề trong thời gian quy định.
13. NGUYÊN TẮC TRỰC QUAN, TOÁN THỰC TẾ VÀ SỰ KHỚP NHAU GIỮA HÌNH VẼ VÀ ĐỀ BÀI: 
    - CHỈ ĐỐI VỚI DUY NHẤT MÔN TOÁN: BẮT BUỘC ở PHẦN TỰ LUẬN phải có ÍT NHẤT 1 BÀI TOÁN TỪ THỰC TẾ (ứng dụng vào đời sống). Bài toán này phải ở mức độ dễ nhưng yêu cầu học sinh phải biết phân tích. ĐẶC BIỆT BẮT BUỘC bài toán này phải đi kèm một hình vẽ (bằng mã SVG) hoặc một hình ảnh minh họa thật rõ nét đính kèm thẳng vào nội dung câu hỏi.
    - ĐỐI VỚI CÁC MÔN HỌC KHÁC (KHÔNG PHẢI TOÁN): TUYỆT ĐỐI KHÔNG ĐƯỢC CHÈN BÀI TOÁN/CÂU HỎI THỰC TẾ VÀO PHẦN TỰ LUẬN.
    - Nếu nội dung trong Phụ lục 3 có chứa dữ liệu biểu đồ, sơ đồ hoặc hình vẽ, BẮT BUỘC đề thi phải có ít nhất 1-2 câu hỏi khai thác các yếu tố này. Đối với biểu đồ (cột, đường, tròn), phải dùng QuickChart.io để tạo ảnh thực tế nhúng vào đề (sử dụng thẻ img với referrerPolicy="no-referrer").
    - TUYỆT ĐỐI BẮT BUỘC (Môn Toán): TẤT CẢ các câu hỏi trắc nghiệm (TNKQ) liên quan đến hình học PHẢI CÓ hình vẽ minh họa đi kèm (sử dụng mã SVG hợp lệ và rõ nét để vẽ hình). 
    - SỰ NHẤT QUÁN: Hình vẽ (SVG) và nội dung văn bản của câu hỏi PHẢI KHỚP NHAU TUYỆT ĐỐI 100%. Ví dụ: Nếu đề bài mô tả "Tam giác ABC vuông tại A", thì hình vẽ SVG BẮT BUỘC phải thể hiện đúng một tam giác vuông tại đỉnh A, có ký hiệu góc vuông, và ghi chú rõ ràng các nhãn A, B, C. Không được sai lệch giữa chữ và hình.
14. ĐỊNH DẠNG TOÁN HỌC NGHIÊM NGẶT: Trình bày TOÀN BỘ công thức và ký hiệu toán học dưới dạng mã LaTeX thuần túy.
    - BẮT BUỘC dùng duy nhất cặp dấu $ cho TẤT CẢ các công thức (bao gồm cả công thức cùng dòng và công thức nằm riêng một dòng). Ví dụ: $\\frac{1}{2}$.
    - TUYỆT ĐỐI KHÔNG dùng cặp dấu \\[ ... \\] hoặc $$ ... $$.
    - TUYỆT ĐỐI KHÔNG dùng ký tự Unicode toán học (ví dụ KHÔNG dùng α, β, ∈, ≠, ≤, ≥, ±, ×, ÷, √ mà PHẢI dùng $\\alpha$, $\\beta$, $\\in$, $\\neq$, $\\le$, $\\ge$, $\\pm$, $\\times$, $\\div$, $\\sqrt{}$).
    - Phân số: TUYỆT ĐỐI không dùng dấu gạch chéo "a/b" hoặc "1/x" trong bất kỳ trường hợp nào (kể cả trong hệ phương trình). BẮT BUỘC phải dùng LaTeX $\\frac{a}{b}$ để hiển thị gạch ngang nằm ngang (ví dụ: $\\frac{1}{x}$).
    - Hệ phương trình/hệ bất phương trình: TUYỆT ĐỐI BẮT BUỘC phải dùng cấu trúc $\\begin{cases} ... \\end{cases}$. Không được dùng dấu ngoặc nhọn kết hợp với text thường.
    - Ký hiệu góc (Angle): TUYỆT ĐỐI KHÔNG DÙNG ký hiệu góc ở phía trước (như $\\angle A$, $\\angle ABC$). BẮT BUỘC phải dùng ký hiệu mũ trên đầu theo chuẩn Toán học Việt Nam: Dùng $\\widehat{ABC}$ cho góc tạo bởi 3 điểm (ví dụ $\\widehat{ABC} = 90^\\circ$) và $\\hat{A}$ cho góc 1 điểm (ví dụ $\\hat{A}$).
    - TUYỆT ĐỐI KHÔNG dùng dấu $ cho các số bình thường, số thập phân, số nguyên không chứa phép toán (ví dụ: viết 125,784 thay vì $125,784$, viết 126 thay vì $126$). Chỉ dùng dấu $ cho các biểu thức, phương trình, ký hiệu toán học hoặc công thức toán học thực sự.
    - ĐẶC BIỆT LƯU Ý: TUYỆT ĐỐI KHÔNG bao gồm văn bản tiếng Việt, chữ cái thông thường, hoặc dấu câu (như dấu chấm, phẩy, dấu ngoặc) bên trong cặp dấu $. 
      + Ví dụ SAI 1: $A, B, C thẳng hàng. Điểm A$ -> Sửa thành ĐÚNG: $A, B, C$ thẳng hàng. Điểm $A$.
      + Ví dụ SAI 2: $AB. c) Điểm A$ -> Sửa thành ĐÚNG: $AB$. c) Điểm $A$.
      + Việc đặt chữ tiếng Việt hoặc dấu câu vào trong cặp dấu $ sẽ gây lỗi font chữ nghiêm trọng khi xuất sang Word và dùng MathType. RÀ SOÁT THẬT KỸ TỪNG CÂU ĐỂ ĐẢM BẢO KHÔNG VI PHẠM LỖI NÀY.
15. TƯƠNG THÍCH MICROSOFT WORD: Đảm bảo không có khoảng trắng thừa giữa dấu $ và nội dung công thức để hỗ trợ sao chép vào Word và chuyển đổi sang Equation dễ dàng.
16. QUY ĐỊNH VỀ MỨC ĐỘ ĐẶC TẢ: Yêu cầu các mức độ (Biết, Hiểu, Vận dụng) của từng bài/câu hỏi trong đề và bảng đặc tả BẮT BUỘC phải lấy chính xác từ tài liệu mức độ đặc tả được cung cấp. Không được tự ý thay đổi mức độ nếu tài liệu đã quy định.
17. YÊU CẦU CẦN ĐẠT: Nội dung cột 'Yêu cầu cần đạt' trong Bảng đặc tả (PHẦN 3) BẮT BUỘC phải dựa trên tài liệu mức độ đặc tả đã tải lên, sau đó TRÌNH BÀY LẠI MỘT CÁCH PHÂN TÁCH VÀ CHI TIẾT THEO TỪNG MỨC ĐỘ phân theo gạch đầu dòng (như yêu cầu số 6). Không được viết thành một đoạn văn hoặc gộp chung. Mức độ nào có câu hỏi (Biết/Hiểu/Vận dụng) thì mới liệt kê yêu cầu cần đạt đó ra.
18. GỘP ĐƠN VỊ KIẾN THỨC: Trong Ma trận (PHẦN 2) và Bảng đặc tả (PHẦN 3), các 'Nội dung/ đơn vị kiến thức' BẮT BUỘC phải được gộp nhóm và trình bày chính xác tuyệt đối theo cấu trúc của tài liệu đặc tả đã tải lên. Không được tự ý chia nhỏ hoặc thay đổi cách phân nhóm các đơn vị kiến thức này.
19. NGUYÊN TẮC TRUNG THỰC TUYỆT ĐỐI: BẮT BUỘC rà soát kỹ lưỡng nội dung Tài liệu mức độ đặc tả đã tải lên. TUYỆT ĐỐI KHÔNG được tự ý thêm thắt, sáng tạo hoặc bổ sung bất kỳ thông tin, yêu cầu cần đạt hay đơn vị kiến thức nào không có trong tài liệu được cung cấp. Mọi nội dung trong đề và bảng đặc tả phải có căn cứ trực tiếp từ tài liệu đầu vào.
20. RÀ SOÁT ĐÚNG CHƯƠNG TRÌNH PHỤ LỤC 3: Bạn BẮT BUỘC phải rà soát đúng và bám sát 100% chương trình học, các đơn vị kiến thức, nội dung bài học có trong file Phụ lục 3 đã tải lên. TUYỆT ĐỐI KHÔNG tự ý sáng tạo, không chèn thêm hoặc bỏ sót bất kỳ chủ đề nào nằm ngoài nội dung của Phụ lục 3. Tất cả nội dung trong Ma trận, Bảng đặc tả và các câu hỏi trong đề thi (PHẦN 4, PHẦN 5) PHẢI nằm gọn trong phạm vi chương trình mà Phụ lục 3 quy định.
21. CẬP NHẬT ĐỊA GIỚI HÀNH CHÍNH MỚI NHẤT (ĐẶC BIỆT ĐỐI VỚI MÔN ĐỊA LÝ): Chú ý NGHIÊM NGẶT đối với các câu hỏi, bài tập, yêu cầu có liên quan đến vị trí địa lý, tên địa danh, xã, huyện, tỉnh, CÁC HUYỆN ĐẢO / THÀNH PHỐ ĐẢO... BẮT BUỘC phải sử dụng đúng tên gọi và cấp đơn vị hành chính theo các Nghị quyết điều chỉnh, sáp nhập địa giới hành chính mới nhất (ví dụ: mô hình chính quyền hai cấp, sáp nhập xã/huyện/tỉnh, thay đổi cách gọi các huyện đảo thành "thành phố đảo" hiện nay như Phú Quốc, không còn đơn vị cũ...). Tuyệt đối không sử dụng tên các địa danh hoặc cấp hành chính cũ đã bị bãi bỏ.
22. YÊU CẦU BẮT BUỘC VỀ THUẬT NGỮ VÀ ĐỊA DANH (CHUẨN CHƯƠNG TRÌNH GDPT 2018 - ÁP DỤNG CHO MỌI MÔN HỌC):
    - ĐỐI VỚI NHÓM MÔN KHOA HỌC TỰ NHIÊN (Toán, Vật lí, Hóa học, Sinh học, KHTN...): Danh pháp khoa học BẮT BUỘC phải sử dụng tên gọi theo chuẩn quốc tế (danh pháp IUPAC, tiếng Anh khoa học). TUYỆT ĐỐI KHÔNG dùng từ phiên âm bồi tiếng Việt. Ví dụ: Dùng Sodium (không dùng Natri), Oxygen (không dùng Oxi), bauxite (không dùng bô-xít), gene (không dùng gen), basalt (không dùng ba-zan).
    - ĐỐI VỚI NHÓM MÔN KHOA HỌC XÃ HỘI (Lịch sử, Địa lí, GDCD...): 
        + Địa danh & Nhân danh quốc tế: Viết nguyên bản theo ngôn ngữ gốc dùng chữ cái Latinh hoặc tiếng Anh. TUYỆT ĐỐI KHÔNG dùng tên phiên âm tiếng Việt. Ví dụ: Dùng Washington, Paris, Australia, Mekong, Christopher Columbus, George Washington. (Ngoại trừ các tên gọi đã Việt hóa hoàn toàn từ lâu đời như Trung Quốc, Nhật Bản, Bắc Kinh...).
        + Địa danh Việt Nam: Sử dụng tên gọi các đơn vị hành chính và phân vùng kinh tế mới nhất hiện hành (Ví dụ: Vùng Bắc Trung Bộ và Duyên hải miền Trung, Thành phố Thủ Đức). Nhất quán sử dụng từ "Biển Đông", "quần đảo Hoàng Sa", "quần đảo Trường Sa".
23. NGUYÊN TẮC CHI TIẾT TUYỆT ĐỐI CỦA ĐỀ THI VÀ HƯỚNG DẪN CHẤM: BẮT BUỘC trình bày toàn bộ nội dung của các đề thi (tất cả các câu hỏi) và hướng dẫn chấm/đáp án một cách THẬT ĐẦY ĐỦ VÀ CHI TIẾT. TUYỆT ĐỐI KHÔNG được dùng các từ ngữ như "tương tự", "tương tự đề 1", "cách làm tương tự", "giống như trên"... để lược bớt hoặc rút gọn nội dung. Đây là nguyên tắc bắt buộc, mọi nội dung, lời giải đều phải được viết rõ ràng, trọn vẹn từ đầu đến cuối cho từng đề riêng biệt.
24. TÁCH BIỆT ĐỀ THI VÀ ĐÁP ÁN: Ở phần đề thi (PHẦN 4, PHẦN 5,...), TUYỆT ĐỐI KHÔNG ghi kèm đáp án, lời giải hay gợi ý vào trong nội dung các câu hỏi của đề thi. Chỉ trình bày nội dung câu hỏi đơn thuần để học sinh làm bài. Toàn bộ đáp án, lời giải chi tiết, hướng dẫn chấm PHẢI được tách riêng và đặt hoàn toàn ở phần "Hướng dẫn chấm/Đáp án".

Trình bày rõ ràng bằng Markdown, sử dụng bảng biểu chuyên nghiệp. Nội dung phải bám sát Phụ lục 3 và Tài liệu mức độ đặc tả được cung cấp.
`;

      const parts: any[] = [{ text: prompt }];
      
      if (fileData) {
        if (fileName.endsWith('.docx') && fileText) {
          parts.push({ text: `Nội dung Phụ lục 3 (trích xuất từ file word): \n${fileText}` });
        } else if (!fileName.endsWith('.docx')) {
          parts.push({
            inlineData: {
              mimeType: mimeType || "application/pdf",
              data: fileData,
            },
          });
        }
      }

      if (specFileData) {
        if (specFileName.endsWith('.docx') && specFileText) {
          parts.push({ text: `Nội dung tài liệu mức độ đặc tả (trích xuất từ file word): \n${specFileText}` });
        } else if (!specFileName.endsWith('.docx')) {
          parts.push({
            inlineData: {
              mimeType: specMimeType || "application/pdf",
              data: specFileData,
            },
          });
        }
        parts.push({ text: "Đây là tài liệu mức độ đặc tả của bảng để tham chiếu mức độ cho từng câu hỏi." });
      }
      
      if (manualInput) {
        parts.push({ text: `Nội dung Phụ lục 3 bổ sung/chỉnh sửa: \n${manualInput}` });
      }

      // Đánh giá dung lượng: Ở ngay trước lúc gọi model, hệ thống đếm số lượng ký tự đầu vào
      let totalInputChars = 0;
      for (const part of parts) {
        if (typeof part.text === "string") {
          totalInputChars += part.text.length;
        } else if (part.inlineData?.data) {
          totalInputChars += part.inlineData.data.length;
        }
      }

      // Ngưỡng > 110.000 ký tự (~32k tokens theo quy định Google Cloud Caching)
      const isLargeInput = totalInputChars > CONTEXT_CACHE_CHAR_THRESHOLD;
      if (isLargeInput) {
        console.log(`[Context Caching Engine] Đầu vào lớn (${totalInputChars.toLocaleString()} ký tự > ${CONTEXT_CACHE_CHAR_THRESHOLD.toLocaleString()}). Kích hoạt cơ chế Context Cache.`);
      }

      let generatedResponseText: string | null = null;
      let lastError: any = null;
      let successfullyUsedModel: string | null = null;
      let usedCacheSuccess = false;

      const ai = new GoogleGenAI({ apiKey: effectiveApiKey });

      // Vòng lặp try-catch qua mảng các model theo thứ tự ưu tiên
      for (const targetModel of MODEL_PRIORITY_LIST) {
        const candidateModels = getModelCandidates(targetModel);
        let candidateSucceeded = false;

        for (const currentModel of candidateModels) {
          try {
            console.log(`[AI Engine] Đang thử kết nối mô hình: ${currentModel}`);
            setLoadingStatus(`Đang kết nối mô hình ${currentModel}...`);

            let cachedContentName: string | null = null;

            // Setup Cache: Nếu Text quá dài (> 110.000 ký tự), gọi caches.create() (TTL 3600s)
            if (isLargeInput) {
              try {
                setLoadingStatus(`Đang nén bộ nhớ đệm Context Cache (TTL: 3600s) cho ${currentModel}...`);
                const cache = await ai.caches.create({
                  model: currentModel,
                  config: {
                    contents: [
                      {
                        role: "user",
                        parts: parts,
                      },
                    ],
                    displayName: `exam_context_cache_${Date.now()}`,
                    ttl: "3600s",
                  },
                });

                if (cache?.name) {
                  cachedContentName = cache.name;
                  console.log(`[Context Caching Engine] Cache thành công với ID: ${cache.name} cho mô hình ${currentModel}.`);
                }
              } catch (cacheErr: any) {
                console.warn(`[Context Caching Engine] Khởi tạo Cache không thành công cho ${currentModel}: ${cacheErr?.message || cacheErr}. Tự động trở về chế độ gửi Prompt đầy đủ theo cách truyền thống.`);
              }
            }

            setLoadingStatus(`Mô hình ${currentModel} đang phân tích và tạo 5 phần hồ sơ...`);

            let response;
            if (cachedContentName) {
              // Rẽ nhánh: Nếu Cache thành công
              console.log(`[AI Engine] Rẽ nhánh Cache: Gửi yêu cầu rút gọn kèm cachedContent tới ${currentModel}.`);
              response = await ai.models.generateContent({
                model: currentModel,
                contents: "Hãy sinh kết quả dựa trên Cache, thực hiện đầy đủ và chính xác 5 PHẦN theo đúng yêu cầu đã lưu trong Cache.",
                config: {
                  cachedContent: cachedContentName,
                },
              });
              usedCacheSuccess = true;
            } else {
              // Rẽ nhánh: Nếu Cache thất bại / Dung lượng file nhỏ
              console.log(`[AI Engine] Rẽ nhánh Truyền thống: Gửi toàn bộ Prompt đầy đủ tới ${currentModel}.`);
              response = await ai.models.generateContent({
                model: currentModel,
                contents: [
                  {
                    role: "user",
                    parts: parts,
                  },
                ],
              });
              usedCacheSuccess = false;
            }

            const resText = response.text;
            if (resText && resText.trim().length > 0) {
              generatedResponseText = resText;
              successfullyUsedModel = currentModel;
              candidateSucceeded = true;
              console.log(`[AI Engine] Thành công tạo nội dung với mô hình: ${currentModel}`);
              break;
            } else {
              throw new Error(`Mô hình ${currentModel} trả về dữ liệu rỗng.`);
            }
          } catch (err: any) {
            lastError = err;
            console.warn(`[AI Engine] Model ${currentModel} báo lỗi (${err?.status || err?.message || err}). Tự động chuyển mượt mà sang thử model tiếp theo mà không gián đoạn trải nghiệm...`);
            setLoadingStatus(`Model ${currentModel} quá tải/bận, tự động chuyển sang model tiếp theo...`);
          }
        }

        if (candidateSucceeded) {
          break;
        }
      }

      if (!generatedResponseText) {
        throw lastError || new Error("Tất cả các mô hình trong danh sách ưu tiên đều không thể tạo nội dung. Vui lòng kiểm tra lại kết nối hoặc khóa API.");
      }

      setActiveModel(successfullyUsedModel);
      setIsCached(usedCacheSuccess);

      const text = generatedResponseText || "Không tạo được nội dung.";

      // Split by "PHẦN" but keep the delimiter or reconstruct
      // Using regex to find the parts more robustly
      const part1Match = text.match(/PHẦN 1[:\s]([\s\S]*?)(?=PHẦN 2|$)/i);
      const part2Match = text.match(/PHẦN 2[:\s]([\s\S]*?)(?=PHẦN 3|$)/i);
      const part3Match = text.match(/PHẦN 3[:\s]([\s\S]*?)(?=PHẦN 4|$)/i);
      const part4Match = text.match(/PHẦN 4[:\s]([\s\S]*?)(?=PHẦN 5|$)/i);
      const part5Match = text.match(/PHẦN 5[:\s]([\s\S]*?$)/i);

      if (part1Match) setLessonPlan(part1Match[1].trim());
      if (part2Match) setMatrix(part2Match[1].trim());
      if (part3Match) setSpecTable(part3Match[1].trim());
      if (part4Match) setExam(part4Match[1].trim());
      if (part5Match) setExam2(part5Match[1].trim());

      // Fallback if regex fails (e.g. model didn't follow exact format)
      if (!part1Match && !part2Match && !part3Match && !part4Match && !part5Match) {
        setLessonPlan(text);
        setError("Mô hình trả về định dạng không chuẩn, hiển thị toàn bộ nội dung ở phần Kế hoạch bài dạy.");
      }

    } catch (err: any) {
      console.error("Error generating content:", err);
      setError(err.message || "Đã xảy ra lỗi khi tạo nội dung. Vui lòng thử lại.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-4 md:p-8">
      <div className="max-w-5xl mx-auto space-y-8">
        
        {/* Header */}
        <header className="text-center space-y-4">
          <div className="space-y-2">
            <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-indigo-50 border border-indigo-200 text-indigo-700 text-xs font-semibold tracking-wide uppercase">
              THCS Quang Trung • GVBM: Lê Tâm
            </div>
            <h1 className="text-3xl md:text-4xl font-bold text-indigo-900 tracking-tight">
              Trợ lý tạo đề kiểm tra theo CV 7991
            </h1>
            <p className="text-slate-600 max-w-2xl mx-auto text-sm md:text-base">
              Hệ thống tự động xây dựng ma trận, bảng đặc tả và đề kiểm tra chuẩn quy định Công văn 7991/BGDĐT-GDTrH.
            </p>
          </div>
          
          <div className="flex items-center justify-center">
            <button
              onClick={() => {
                setTempApiKey(apiKey);
                setShowApiKeyModal(true);
              }}
              className={`flex items-center gap-2.5 px-5 py-3 rounded-2xl text-sm font-semibold border shadow-sm transition-all hover:scale-[1.02] active:scale-95 ${
                apiKey
                  ? "bg-emerald-50 text-emerald-800 border-emerald-300 hover:bg-emerald-100"
                  : "bg-amber-50 text-amber-800 border-amber-300 hover:bg-amber-100 ring-2 ring-amber-400/30 animate-pulse"
              }`}
            >
              <Key className="w-5 h-5 text-indigo-600" />
              <div className="text-left">
                <div className="text-[11px] text-slate-500 font-normal leading-none">Cấu hình Google AI</div>
                <div className="font-bold">{apiKey ? "API Key: Đã kết nối ✓" : "Nhập Gemini API Key"}</div>
              </div>
            </button>
          </div>
        </header>

        {/* Modal Cấu hình API Key */}
        <AnimatePresence>
          {showApiKeyModal && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="bg-white rounded-2xl shadow-2xl border border-slate-200 p-6 max-w-md w-full space-y-4"
              >
                <div className="flex justify-between items-center">
                  <div className="flex items-center gap-2 text-indigo-900 font-bold text-lg">
                    <Key className="w-5 h-5 text-indigo-600" />
                    <span>Cấu hình Gemini API Key</span>
                  </div>
                  <button
                    onClick={() => setShowApiKeyModal(false)}
                    className="p-1 rounded-lg hover:bg-slate-100 text-slate-500"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
                <p className="text-sm text-slate-600 leading-relaxed">
                  Nhập khóa API Google Gemini để hệ thống phân tích phụ lục và tạo đề thi theo Công văn 7991. Khóa được lưu an toàn trực tiếp trên trình duyệt của bạn.
                </p>
                <div className="space-y-2">
                  <label className="text-xs font-semibold text-slate-700">Gemini API Key:</label>
                  <input
                    type="password"
                    value={tempApiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    placeholder="Dán mã API Key (AIzaSy...)"
                    className="w-full p-3 rounded-xl border border-slate-200 bg-slate-50 text-sm font-mono outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                  <a
                    href="https://aistudio.google.com/app/apikey"
                    target="_blank"
                    rel="noreferrer"
                    className="text-xs text-indigo-600 hover:text-indigo-800 flex items-center gap-1 mt-1 font-medium"
                  >
                    <span>Lấy API Key miễn phí tại Google AI Studio</span>
                    <ExternalLink className="w-3 h-3" />
                  </a>
                </div>
                <div className="flex justify-end gap-3 pt-2">
                  <button
                    onClick={() => setShowApiKeyModal(false)}
                    className="px-4 py-2 text-sm text-slate-600 hover:bg-slate-100 rounded-xl"
                  >
                    Đóng
                  </button>
                  <button
                    onClick={handleSaveApiKey}
                    className="px-5 py-2 text-sm bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-xl shadow-sm"
                  >
                    Lưu cấu hình
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Upload Section */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8 space-y-8"
        >
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <div className="space-y-6">
              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-700">Chọn loại đề kiểm tra:</label>
                <select 
                  value={examType}
                  onChange={(e) => setExamType(e.target.value)}
                  className="w-full p-3 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                >
                  {examTypes.map(type => (
                    <option key={type} value={type}>{type}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-700">Thời gian làm bài (phút):</label>
                <input 
                  type="number"
                  value={duration}
                  onChange={(e) => setDuration(parseInt(e.target.value) || 0)}
                  className="w-full p-3 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                />
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-700">Nội dung Phụ lục 3 (Dán trực tiếp hoặc chỉnh sửa):</label>
                <textarea
                  value={manualInput}
                  onChange={(e) => setManualInput(e.target.value)}
                  placeholder="Dán nội dung Phụ lục 3 vào đây hoặc chỉnh sửa sau khi tải file..."
                  className="w-full h-48 p-3 rounded-xl border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500 resize-none"
                />
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-700">Tải lên Phụ lục 3:</label>
                <div 
                  onClick={triggerFileInput}
                  className={`border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-colors h-[52px] flex items-center justify-center gap-2 ${
                    fileName ? "border-green-500 bg-green-50" : "border-slate-300 hover:border-indigo-400 hover:bg-slate-50"
                  }`}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".pdf,.doc,.docx,.txt,.md"
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                  {fileName ? (
                    <>
                      <CheckCircle className="w-5 h-5 text-green-500" />
                      <span className="text-sm font-medium text-green-700 truncate max-w-[200px]">{fileName}</span>
                    </>
                  ) : (
                    <>
                      <Upload className="w-5 h-5 text-slate-400" />
                      <span className="text-sm text-slate-500">Chọn file Phụ lục 3</span>
                    </>
                  )}
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-700">Tải lên Mức độ đặc tả (Tùy chọn):</label>
                <div 
                  onClick={triggerSpecFileInput}
                  className={`border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-colors h-[52px] flex items-center justify-center gap-2 ${
                    specFileName ? "border-green-500 bg-green-50" : "border-slate-300 hover:border-indigo-400 hover:bg-slate-50"
                  }`}
                >
                  <input
                    ref={specFileInputRef}
                    type="file"
                    accept=".pdf,.doc,.docx,.txt,.md"
                    onChange={handleSpecFileUpload}
                    className="hidden"
                  />
                  {specFileName ? (
                    <>
                      <CheckCircle className="w-5 h-5 text-green-500" />
                      <span className="text-sm font-medium text-green-700 truncate max-w-[200px]">{specFileName}</span>
                    </>
                  ) : (
                    <>
                      <Upload className="w-5 h-5 text-slate-400" />
                      <span className="text-sm text-slate-500">Chọn file Mức độ đặc tả</span>
                    </>
                  )}
                </div>
              </div>
            </div>

            <div className="space-y-4">
              <label className="text-sm font-semibold text-slate-700">Tỉ lệ điểm (Biết - Hiểu - Vận dụng):</label>
              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Biết (%)</span>
                  <input 
                    type="number" 
                    value={ratios.know}
                    onChange={(e) => setRatios({...ratios, know: parseInt(e.target.value) || 0})}
                    className="w-full p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Hiểu (%)</span>
                  <input 
                    type="number" 
                    value={ratios.understand}
                    onChange={(e) => setRatios({...ratios, understand: parseInt(e.target.value) || 0})}
                    className="w-full p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Vận dụng (%)</span>
                  <input 
                    type="number" 
                    value={ratios.apply}
                    onChange={(e) => setRatios({...ratios, apply: parseInt(e.target.value) || 0})}
                    className="w-full p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>
              
              <label className="text-sm font-semibold text-slate-700 block pt-2">Số lượng câu hỏi và Điểm số:</label>
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Trắc nghiệm (1/4)</span>
                  <div className="flex gap-2">
                    <input 
                      type="number" 
                      value={counts.multipleChoice}
                      onChange={(e) => setCounts({...counts, multipleChoice: parseInt(e.target.value) || 0})}
                      placeholder="Số câu"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                    <input 
                      type="number" 
                      step="0.25"
                      value={points.multipleChoice}
                      onChange={(e) => setPoints({...points, multipleChoice: parseFloat(e.target.value) || 0})}
                      placeholder="Điểm"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Đúng / Sai</span>
                  <div className="flex gap-2">
                    <input 
                      type="number" 
                      value={counts.trueFalse}
                      onChange={(e) => {
                        const val = parseInt(e.target.value) || 0;
                        if (val === 0 && points.trueFalse > 0) {
                          setPoints(prev => ({...prev, multipleChoice: prev.multipleChoice + prev.trueFalse, trueFalse: 0}));
                        }
                        setCounts({...counts, trueFalse: val});
                      }}
                      placeholder="Số câu"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                    <input 
                      type="number" 
                      step="0.25"
                      value={points.trueFalse}
                      onChange={(e) => setPoints({...points, trueFalse: parseFloat(e.target.value) || 0})}
                      placeholder="Điểm"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Trả lời ngắn</span>
                  <div className="flex gap-2">
                    <input 
                      type="number" 
                      value={counts.shortAnswer}
                      onChange={(e) => {
                        const val = parseInt(e.target.value) || 0;
                        if (val === 0 && points.shortAnswer > 0) {
                          setPoints(prev => ({...prev, multipleChoice: prev.multipleChoice + prev.shortAnswer, shortAnswer: 0}));
                        }
                        setCounts({...counts, shortAnswer: val});
                      }}
                      placeholder="Số câu"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                    <input 
                      type="number" 
                      step="0.25"
                      value={points.shortAnswer}
                      onChange={(e) => setPoints({...points, shortAnswer: parseFloat(e.target.value) || 0})}
                      placeholder="Điểm"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <span className="text-xs text-slate-500">Tự luận</span>
                  <div className="flex gap-2">
                    <input 
                      type="number" 
                      value={counts.essay}
                      onChange={(e) => setCounts({...counts, essay: parseInt(e.target.value) || 0})}
                      placeholder="Số câu"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                    <input 
                      type="number" 
                      step="0.25"
                      value={points.essay}
                      onChange={(e) => setPoints({...points, essay: parseFloat(e.target.value) || 0})}
                      placeholder="Điểm"
                      className="w-1/2 p-2 rounded-lg border border-slate-200 bg-slate-50 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>

          {error && (
            <div className="p-4 bg-red-50 border border-red-200 rounded-xl flex items-start gap-3">
              <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
              <div className="text-red-700 text-sm space-y-1">
                {error.includes('•') && <p className="font-semibold mb-2">Vui lòng kiểm tra lại các thông tin sau:</p>}
                {error.split('\n').map((err, idx) => (
                  <p key={idx}>{err}</p>
                ))}
              </div>
            </div>
          )}

          <div className="flex justify-center pt-2">
            <button
              onClick={generateAll}
              disabled={loading}
              className={`flex items-center gap-2 px-8 py-3 rounded-xl font-semibold text-white shadow-md transition-all ${
                loading
                  ? "bg-slate-400 cursor-not-allowed"
                  : "bg-indigo-600 hover:bg-indigo-700 hover:shadow-lg active:scale-95"
              }`}
            >
              {loading ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  <span>{loadingStatus}</span>
                </>
              ) : (
                <>
                  <RefreshCw className="w-5 h-5" />
                  Sinh Ma Trận - Đặc Tả - Đề
                </>
              )}
            </button>
          </div>

          {loading && (
            <div className="flex items-center justify-center gap-2 text-xs text-indigo-600 font-medium pt-1">
              <span className="inline-block w-2 h-2 rounded-full bg-indigo-500 animate-pulse" />
              <span>Hệ thống tự động điều phối chuỗi model ưu tiên và kiểm tra bộ nhớ đệm Context Cache...</span>
            </div>
          )}
        </motion.div>

        {/* Results Section */}
        <AnimatePresence>
          {(lessonPlan || matrix || specTable || exam || exam2) && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-8"
            >
              <div className="flex flex-col sm:flex-row items-center justify-between gap-4 bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                <div className="flex items-center gap-3">
                  <div className="w-9 h-9 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-600 shrink-0">
                    <CheckCircle className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Mô hình AI hoàn thành:</div>
                    <div className="text-sm font-bold text-slate-800 flex items-center gap-2">
                      <span>{activeModel || "Tối ưu hóa đa mô hình"}</span>
                      {isCached && (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-full">
                          ⚡ Đã tối ưu Context Cache
                        </span>
                      )}
                    </div>
                  </div>
                </div>

                <button
                  onClick={handleExportWordAll}
                  className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 text-white rounded-xl font-medium hover:bg-blue-700 transition-colors shadow-sm w-full sm:w-auto justify-center"
                >
                  <FileDown className="w-5 h-5" />
                  Xuất toàn bộ Hồ sơ (Word)
                </button>
              </div>

              <div ref={resultsRef} className="space-y-8">
                {/* Lesson Plan Section */}
                {lessonPlan && (
                  <ResultCard title="PHẦN 1: KẾ HOẠCH BÀI DẠY" content={lessonPlan} onUpdate={setLessonPlan} color="orange" />
                )}

                {/* Matrix Section */}
                {matrix && (
                  <ResultCard title="PHẦN 2: MA TRẬN ĐỀ" content={matrix} onUpdate={setMatrix} color="blue" />
                )}

                {/* Spec Table Section */}
                {specTable && (
                  <ResultCard title="PHẦN 3: BẢNG ĐẶC TẢ" content={specTable} onUpdate={setSpecTable} color="purple" />
                )}

                {/* Exam Section */}
                {exam && (
                  <ResultCard title="PHẦN 4: ĐỀ KIỂM TRA SỐ 1" content={exam} onUpdate={setExam} color="emerald" />
                )}

                {/* Exam 2 Section */}
                {exam2 && (
                  <ResultCard title="PHẦN 5: ĐỀ KIỂM TRA SỐ 2" content={exam2} onUpdate={setExam2} color="emerald" />
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}

const transformSvgProps = (props: any) => {
  const { node, ...rest } = props;
  const newProps: any = {};
  for (const key in rest) {
    if (key.includes('-') && !key.startsWith('data-') && !key.startsWith('aria-')) {
      const camelKey = key.replace(/-([a-z])/g, (g) => g[1].toUpperCase());
      newProps[camelKey] = rest[key];
    } else if (key === 'class') {
      newProps.className = rest[key];
    } else if (key === 'viewbox') {
      newProps.viewBox = rest[key];
    } else if (key === 'xlink:href') {
      newProps.xlinkHref = rest[key];
    } else {
      newProps[key] = rest[key];
    }
  }
  return newProps;
};

const StandaloneSvg = ({ children, width, height, viewBox, style }: any) => (
  <SvgContext.Provider value={true}>
    <svg
      width={width}
      height={height}
      viewBox={viewBox}
      className="inline-block align-middle my-1 overflow-visible"
      style={{ maxWidth: '100%', ...style }}
    >
      {children}
    </svg>
  </SvgContext.Provider>
);

function ResultCard({ title, content, onUpdate, color }: { title: string, content: string, onUpdate: (val: string) => void, color: "blue" | "purple" | "emerald" | "orange" }) {
  const [copied, setCopied] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [editContent, setEditContent] = useState(content);
  const containerRef = useRef<HTMLDivElement>(null);

  const handleCopy = () => {
    navigator.clipboard.writeText(content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleSave = () => {
    onUpdate(editContent);
    setIsEditing(false);
  };

  const handleExportExcel = () => {
    if (!containerRef.current) return;
    
    const table = containerRef.current.querySelector("table");
    if (!table) {
      alert("Không tìm thấy bảng để xuất.");
      return;
    }

    const wb = XLSX.utils.table_to_book(table, { raw: true });
    XLSX.writeFile(wb, `${title.replace(/[:/\\?%*|"<>]/g, "_")}.xlsx`);
  };

  const handleExportWord = async () => {
    if (!containerRef.current) return;
    
    const processedHtml = await processHtmlForDocx(containerRef.current);
    
    const header = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title>
    <style>
      body { 
        font-family: 'Times New Roman', Times, serif; 
        font-size: 12pt; 
        mso-ascii-font-family: 'Times New Roman'; 
        mso-hansi-font-family: 'Times New Roman'; 
        mso-bidi-font-family: 'Times New Roman';
      }
      table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
      th, td { border: 1px solid black; padding: 8px; text-align: left; font-family: 'Times New Roman', Times, serif; }
      th { background-color: #f2f2f2; font-weight: bold; text-align: center; }
      h1, h2, h3, h4, h5, h6 { 
        font-family: 'Times New Roman', Times, serif; 
        mso-ascii-font-family: 'Times New Roman'; 
        mso-hansi-font-family: 'Times New Roman'; 
      }
      p { line-height: 1.5; margin-bottom: 10px; font-family: 'Times New Roman', Times, serif; }
      img { max-width: 100%; height: auto; }
    </style>
    </head><body>`;
    const footer = "</body></html>";
    
    const mathTypeNote = `<p style="color: red; font-weight: bold; margin-bottom: 20px; font-style: italic;">
      Hãy sử dụng MathType để chuyển các công thức Toán học, Hóa học. Cách làm cụ thể:<br/>
      Bôi đen cả dòng chứa công thức (thường có dấu $), bấm MathType, chọn Toogle Text
    </p>`;
    
    const htmlContent = mathTypeNote + `<h2 style="text-align: center; color: #1e3a8a; font-size: 16pt; text-transform: uppercase; margin-bottom: 20px;">${title}</h2>` + processedHtml;
    const sourceHTML = header + htmlContent + footer;
    
    try {
      const blob = await asBlob(sourceHTML) as Blob;
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${title.replace(/[:/\\?%*|"<>]/g, "_")}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Export to DOCX failed:", error);
      alert("Đã xảy ra lỗi khi xuất file Word. Vui lòng thử lại.");
    }
  };

  const hasTable = content.toLowerCase().includes("<table") || content.toLowerCase().includes("|---|");

  const colorClasses = {
    blue: "border-blue-200 bg-blue-50 text-blue-900",
    purple: "border-purple-200 bg-purple-50 text-purple-900",
    emerald: "border-emerald-200 bg-emerald-50 text-emerald-900",
    orange: "border-orange-200 bg-orange-50 text-orange-900",
  };

  const headerColors = {
    blue: "text-blue-700",
    purple: "text-purple-700",
    emerald: "text-emerald-700",
    orange: "text-orange-700",
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden"
    >
      <div className={`px-6 py-4 border-b ${colorClasses[color]} bg-opacity-30 flex justify-between items-center`}>
        <h2 className={`text-xl font-bold ${headerColors[color]} flex items-center gap-2`}>
          <FileText className="w-5 h-5" />
          {title}
        </h2>
        <div className="flex items-center gap-2">
          <button
            onClick={handleExportWord}
            className="p-2 rounded-lg transition-colors hover:bg-white/50 text-slate-600"
            title="Xuất sang Word"
          >
            <FileDown className="w-5 h-5" />
          </button>
          {hasTable && (
            <button
              onClick={handleExportExcel}
              className="p-2 rounded-lg transition-colors hover:bg-white/50 text-slate-600"
              title="Xuất sang Excel/Trang tính"
            >
              <FileSpreadsheet className="w-5 h-5" />
            </button>
          )}
          <button
            onClick={() => {
              if (isEditing) handleSave();
              else setIsEditing(true);
            }}
            className={`p-2 rounded-lg transition-colors ${
              isEditing ? "bg-indigo-100 text-indigo-700" : "hover:bg-white/50 text-slate-600"
            }`}
            title={isEditing ? "Lưu chỉnh sửa" : "Chỉnh sửa nội dung"}
          >
            {isEditing ? <Save className="w-5 h-5" /> : <Edit2 className="w-5 h-5" />}
          </button>
          <button
            onClick={handleCopy}
            className={`p-2 rounded-lg transition-colors ${
              copied ? "bg-green-100 text-green-700" : "hover:bg-white/50 text-slate-600"
            }`}
            title="Sao chép nội dung"
          >
            {copied ? <Check className="w-5 h-5" /> : <Copy className="w-5 h-5" />}
          </button>
        </div>
      </div>
      <div ref={containerRef} className="p-6 overflow-x-auto prose prose-slate max-w-none">
        {isEditing ? (
          <textarea
            value={editContent}
            onChange={(e) => setEditContent(e.target.value)}
            className="w-full h-[500px] p-4 rounded-xl border border-slate-200 bg-slate-50 font-mono text-sm outline-none focus:ring-2 focus:ring-indigo-500"
          />
        ) : (
          <ReactMarkdown 
            remarkPlugins={[remarkGfm, remarkMath]} 
            rehypePlugins={[rehypeRaw, rehypeKatex]}
            components={{
              svg: (props) => (
                <SvgContext.Provider value={true}>
                  <svg {...transformSvgProps(props)} />
                </SvgContext.Provider>
              ),
              line: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <line {...transformSvgProps(props)} />;
                const hasSvgAttr = props.x1 !== undefined || props.x2 !== undefined || props.y1 !== undefined || props.y2 !== undefined || props.stroke !== undefined;
                if (hasSvgAttr) {
                  const x1 = parseFloat(props.x1) || 0;
                  const x2 = parseFloat(props.x2) || 100;
                  const y1 = parseFloat(props.y1) || 0;
                  const y2 = parseFloat(props.y2) || 100;
                  const minX = Math.min(x1, x2);
                  const minY = Math.min(y1, y2);
                  const maxX = Math.max(x1, x2, minX + 10);
                  const maxY = Math.max(y1, y2, minY + 10);
                  const w = Math.max(maxX - minX + 10, 10);
                  const h = Math.max(maxY - minY + 10, 10);
                  return (
                    <StandaloneSvg width={w} height={h} viewBox={`${minX - 5} ${minY - 5} ${w} ${h}`}>
                      <line {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return props.children ? <span className="block my-1" {...transformSvgProps(props)} /> : <hr className="my-2 border-slate-300" />;
              },
              rect: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <rect {...transformSvgProps(props)} />;
                const hasSvgAttr = props.width !== undefined || props.height !== undefined || props.fill !== undefined || props.stroke !== undefined || props.x !== undefined || props.y !== undefined;
                if (hasSvgAttr) {
                  const w = parseFloat(props.width) || 100;
                  const h = parseFloat(props.height) || 100;
                  return (
                    <StandaloneSvg width={w} height={h} viewBox={`0 0 ${w} ${h}`}>
                      <rect {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <div className="border border-slate-300 p-2 my-1 rounded" {...transformSvgProps(props)} />;
              },
              ellipse: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <ellipse {...transformSvgProps(props)} />;
                const hasSvgAttr = props.cx !== undefined || props.cy !== undefined || props.rx !== undefined || props.ry !== undefined || props.fill !== undefined || props.stroke !== undefined;
                if (hasSvgAttr) {
                  const rx = parseFloat(props.rx) || 50;
                  const ry = parseFloat(props.ry) || 30;
                  const cx = parseFloat(props.cx) || rx;
                  const cy = parseFloat(props.cy) || ry;
                  const w = Math.max(cx + rx, rx * 2);
                  const h = Math.max(cy + ry, ry * 2);
                  return (
                    <StandaloneSvg width={w} height={h} viewBox={`0 0 ${w} ${h}`}>
                      <ellipse {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <span className="inline-block border border-current rounded-full px-2 py-0.5 mx-1" {...transformSvgProps(props)} />;
              },
              circle: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <circle {...transformSvgProps(props)} />;
                const hasSvgAttr = props.cx !== undefined || props.cy !== undefined || props.r !== undefined || props.fill !== undefined || props.stroke !== undefined;
                if (hasSvgAttr) {
                  const r = parseFloat(props.r) || 20;
                  const cx = parseFloat(props.cx) || r;
                  const cy = parseFloat(props.cy) || r;
                  const size = Math.max(cx + r, r * 2);
                  return (
                    <StandaloneSvg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
                      <circle {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <span className="inline-block border border-current rounded-full px-1 mx-1" {...transformSvgProps(props)} />;
              },
              path: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <path {...transformSvgProps(props)} />;
                if (props.d) {
                  return (
                    <StandaloneSvg>
                      <path {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <span {...transformSvgProps(props)} />;
              },
              g: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <g {...transformSvgProps(props)} />;
                return (
                  <StandaloneSvg>
                    <g {...transformSvgProps(props)} />
                  </StandaloneSvg>
                );
              },
              polygon: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <polygon {...transformSvgProps(props)} />;
                if (props.points) {
                  return (
                    <StandaloneSvg>
                      <polygon {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <span {...transformSvgProps(props)} />;
              },
              polyline: (props: any) => {
                const isSvg = useContext(SvgContext);
                if (isSvg) return <polyline {...transformSvgProps(props)} />;
                if (props.points) {
                  return (
                    <StandaloneSvg>
                      <polyline {...transformSvgProps(props)} />
                    </StandaloneSvg>
                  );
                }
                return <span {...transformSvgProps(props)} />;
              },
              text: (props: any) => {
                const isSvg = useContext(SvgContext);
                return isSvg ? <text {...transformSvgProps(props)} /> : <span {...transformSvgProps(props)} />;
              }
            }}
          >
            {content}
          </ReactMarkdown>
        )}
      </div>
    </motion.div>
  );
}
