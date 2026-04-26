import React, { useState, useCallback, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
// @ts-ignore
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import { PDFDocument } from 'pdf-lib';
import ExcelJS from 'exceljs';
import { 
  FileSpreadsheet, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  Download, 
  RefreshCcw, 
  LayoutList,
  AlertTriangle,
  Loader2,
  Table as TableIcon,
  HelpCircle,
  Search,
  Eye,
  ArrowRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import { OrderInfo, PDFPageInfo, MatchResult, ProcessingStats } from './types';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export default function App() {
  // Files
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  
  // Data
  const [excelOrders, setExcelOrders] = useState<OrderInfo[]>([]);
  const [pdfPages, setPdfPages] = useState<PDFPageInfo[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState<string>('');

  // States
  const [activeTab, setActiveTab] = useState<'pdf' | 'excel'>('pdf');
  const [step, setStep] = useState(1); // 1: Upload, 2: Preview/Processing

  // Excel Refiner Files/State
  const [rawExcelFile, setRawExcelFile] = useState<File | null>(null);
  const [excelMapping, setExcelMapping] = useState<Record<string, string>>({});
  const [rawExcelData, setRawExcelData] = useState<any[]>([]);

  const excelRefinerInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);

  const resetAll = () => {
    setExcelFile(null);
    setPdfFile(null);
    setExcelOrders([]);
    setPdfPages([]);
    setIsProcessing(false);
    setProgress(0);
    setError(null);
    setStep(1);
    setStatusMessage('');
    setRawExcelFile(null);
    setExcelMapping({});
    setRawExcelData([]);
    if (excelInputRef.current) excelInputRef.current.value = '';
    if (pdfInputRef.current) pdfInputRef.current.value = '';
    if (excelRefinerInputRef.current) excelRefinerInputRef.current.value = '';
  };

  const handleRawExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 100 * 1024 * 1024) {
      setError("File Excel không được vượt quá 100MB.");
      return;
    }

    setRawExcelFile(file);
    setError(null);
    setStatusMessage("Đang phân tích file...");

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json<any>(sheet, { header: 'A', raw: false });

      if (jsonData.length === 0) throw new Error("File trống.");

      // Find header row
      let headerIdx = -1;
      let targetMapping: Record<string, string> = {};
      
      const targets = [
        { label: "STT", keywords: ["stt", "số thứ tự", "no.", "index"] },
        { label: "Mã đơn hàng", keywords: ["đơn hàng", "order id", "mã đơn"] },
        { label: "Tên sản phẩm", keywords: ["tên sản phẩm", "product name", "tên hàng"] },
        { label: "Mã SKU", keywords: ["sku", "mã hàng", "phân loại", "variant"] },
        { 
          label: "Số lượng sản phẩm", 
          keywords: ["số lượng", "quantity", "qty"],
          exclude: ["tổng", "total"] 
        },
        { label: "Giá sản phẩm", keywords: ["giá", "price", "thành tiền", "đơn giá"] },
        { label: "Ghi chú", keywords: ["ghi chú", "note"] }
      ];

      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i];
        let foundCount = 0;
        const currentMapping: Record<string, string> = {};

        for (const target of targets) {
          for (const [key, val] of Object.entries(row)) {
            const lowVal = String(val || '').toLowerCase().trim();
            
            const matches = target.keywords.some(k => lowVal.includes(k));
            const excluded = ('exclude' in target) && (target.exclude as string[]).some(e => lowVal.includes(e));

            if (matches && !excluded) {
              currentMapping[target.label] = key;
              foundCount++;
              break;
            }
          }
        }

        if (foundCount >= 2) {
          headerIdx = i;
          targetMapping = currentMapping;
          break;
        }
      }

      if (headerIdx === -1) {
        throw new Error("Không thể tự động nhận diện các cột. Vui lòng kiểm tra lại file Excel.");
      }

      setExcelMapping(targetMapping);
      setRawExcelData(jsonData.slice(headerIdx + 1));
      setStatusMessage("Phân tích xong. Sẵn sàng tải xuống.");

    } catch (err: any) {
      setError(`Lỗi: ${err.message}`);
      setRawExcelFile(null);
    }
  };

  const exportCleanedExcel = async () => {
    if (!rawExcelData.length) return;
    
    setIsProcessing(true);
    setStatusMessage("Đang tạo file excel với 2 trang...");
    try {
      const workbook = new ExcelJS.Workbook();
      
      // --- SHEET 1: CHI TIẾT ---
      const worksheet1 = workbook.addWorksheet('Đơn hàng');
      const headers1 = ["STT", "Mã đơn hàng", "Tên sản phẩm", "Mã SKU", "Số lượng sản phẩm", "Giá sản phẩm", "Ghi chú"];
      
      const headerRow1 = worksheet1.addRow(headers1);
      headerRow1.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        // Headers left by default, right for specified numeric columns
        if (colNumber === 5 || colNumber === 6) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      let lastSttValue: any = null;
      let mergeStartRow = 2;

      rawExcelData.forEach((row, idx) => {
        // Carry forward STT if empty (common for merged cells in source)
        const rawStt = row[excelMapping["STT"]];
        const currentStt = (rawStt !== undefined && rawStt !== null && String(rawStt).trim() !== "") 
          ? String(rawStt).trim() 
          : lastSttValue;

        const rowData = [
          currentStt,
          row[excelMapping["Mã đơn hàng"]] || "",
          row[excelMapping["Tên sản phẩm"]] || "",
          row[excelMapping["Mã SKU"]] || "",
          Number(row[excelMapping["Số lượng sản phẩm"]] || 0),
          row[excelMapping["Giá sản phẩm"]] || "",
          row[excelMapping["Ghi chú"]] || ""
        ];
        
        const newRow = worksheet1.addRow(rowData);
        const currentRowNum = newRow.number;

        // Apply borders and alignment
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          // Numeric columns right, STT center, rest left
          if (colNumber === 1) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else if (colNumber === 5 || colNumber === 6) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });

        // Merge STT Column (Col A) logic
        if (idx > 0 && currentStt === lastSttValue && currentStt !== null) {
          // Still in the same STT group
        } else {
          if (idx > 0 && currentRowNum - 1 > mergeStartRow) {
            worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum - 1}`);
          }
          mergeStartRow = currentRowNum;
        }
        
        lastSttValue = currentStt;

        // Final merge for the last group
        if (idx === rawExcelData.length - 1 && currentRowNum > mergeStartRow) {
          worksheet1.mergeCells(`A${mergeStartRow}:A${currentRowNum}`);
        }
      });

      worksheet1.columns.forEach((column, index) => {
        if (index === 0) column.width = 8; // STT
        else if (index === 1) column.width = 25; // Order ID
        else if (index === 2) column.width = 40; // Product Name
        else column.width = 20;
      });

      // --- SHEET 2: TỔNG HỢP ---
      const worksheet2 = workbook.addWorksheet('Danh sách sản phẩm');
      
      const summaryMap = new Map<string, { name: string, sku: string, totalQty: number }>();
      
      rawExcelData.forEach(row => {
        const sku = String(row[excelMapping["Mã SKU"]] || 'N/A').trim();
        const name = String(row[excelMapping["Tên sản phẩm"]] || 'N/A').trim();
        const qty = Number(row[excelMapping["Số lượng sản phẩm"]] || 0);
        
        const key = `${sku}|${name}`;
        if (summaryMap.has(key)) {
          summaryMap.get(key)!.totalQty += qty;
        } else {
          summaryMap.set(key, { name, sku, totalQty: qty });
        }
      });

      const headers2 = ["Tên sản phẩm", "Mã SKU", "Sum of Số lượng sản phẩm"];
      const headerRow2 = worksheet2.addRow(headers2);
      headerRow2.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        // Headers left as requested (all headers left, except specific ones in sheet 1)
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      const sortedSummary = Array.from(summaryMap.values()).sort((a, b) => a.name.localeCompare(b.name));

      sortedSummary.forEach((data) => {
        const newRow = worksheet2.addRow([data.name, data.sku, data.totalQty]);
        newRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
          };
          // Column 1 bold, Column 3 right-aligned
          if (colNumber === 1) {
            cell.font = { bold: true };
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          } else if (colNumber === 3) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left' };
          }
        });
      });

      // Add AutoFilter for Columns 1 and 2
      worksheet2.autoFilter = {
        from: 'A1',
        to: 'B' + (sortedSummary.length + 1)
      };

      const totalRow = worksheet2.addRow(['Grand Total', '', sortedSummary.reduce((acc, curr) => acc + curr.totalQty, 0)]);
      totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        if (colNumber === 3) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });

      worksheet2.columns.forEach((column, index) => {
        column.width = index === 0 ? 50 : 25;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `orders_refined_${new Date().getTime()}.xlsx`;
      link.click();
      
      setStatusMessage("Đã xuất hoàn tất file gồm 2 trang.");
    } catch (err: any) {
      setError(`Lỗi xuất file: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const cleanOrderId = (val: any): string => {
    return String(val || '').trim().replace(/['"\s]/g, '');
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // 100MB check
    if (file.size > 100 * 1024 * 1024) {
      setError("File Excel/CSV không được vượt quá 100MB.");
      return;
    }

    setExcelFile(file);
    setError(null);
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      
      let sheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'orders');
      if (!sheetName && workbook.SheetNames.length > 0) {
        sheetName = workbook.SheetNames[0];
      }

      if (!sheetName) throw new Error("Không thể tìm thấy sheet \"Orders\".");

      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 'A', raw: false });

      let headerRowIndex = -1;
      let orderIdColumnKey = 'B';

      // Advanced header detection
      for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
        const row = jsonData[i];
        for (const key in row) {
          const val = String(row[key] || '').toLowerCase().trim();
          if (val === 'mã đơn hàng' || val === 'order id' || val === 'mã đơn' || val === 'orderid') {
            headerRowIndex = i;
            orderIdColumnKey = key;
            break;
          }
        }
        if (headerRowIndex !== -1) break;
      }

      const orders: OrderInfo[] = [];
      const seenOrders = new Set<string>();

      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        let orderId = cleanOrderId(row[orderIdColumnKey]);
        
        if (!orderId || orderId === 'undefined' || orderId === '') continue;

        if (!seenOrders.has(orderId)) {
          seenOrders.add(orderId);
          orders.push({
            stt: orders.length + 1,
            orderId,
            sourceRow: i + 1
          });
        }
      }

      if (orders.length === 0) {
        throw new Error("Không tìm thấy mã đơn hàng nào trong file Excel.");
      }

      setExcelOrders(orders);
    } catch (err: any) {
      setError(`Lỗi Excel: ${err.message}`);
      setExcelFile(null);
      setExcelOrders([]);
    }
  };

  const handlePdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // 100MB check
    if (file.size > 100 * 1024 * 1024) {
      setError("File PDF không được vượt quá 100MB.");
      return;
    }

    setPdfFile(file);
    setError(null);
  };

  const processFiles = async () => {
    if (!excelFile || !pdfFile || excelOrders.length === 0) return;
    
    setIsProcessing(true);
    setProgress(0);
    setError(null);
    setStatusMessage('Đang trích xuất dữ liệu từ PDF...');
    
    try {
      const arrayBuffer = await pdfFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const totalPages = pdf.numPages;
      const pagesInfo: PDFPageInfo[] = [];

      const orderIdRegex = /Order\s*ID\s*:\s*(\d{10,35})/i;
      const excelOrderIds = excelOrders.map(o => o.orderId);

      for (let i = 1; i <= totalPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const strItems = textContent.items.map((item: any) => item.str);
        const fullText = strItems.join(' ');
        const compressedText = strItems.join('').replace(/\s/g, '');
        
        let foundOrderId: string | null = null;
        
        const match = fullText.match(orderIdRegex);
        if (match && match[1]) {
          foundOrderId = match[1];
        } else {
          for (const id of excelOrderIds) {
            if (fullText.includes(id) || compressedText.includes(id)) {
              foundOrderId = id;
              break;
            }
          }
        }

        pagesInfo.push({
          pageIndex: i - 1,
          orderId: foundOrderId,
          status: foundOrderId ? 'matched' : 'not_found'
        });

        setProgress(Math.round((i / totalPages) * 100));
      }

      setPdfPages(pagesInfo);
      setStep(2);
      setStatusMessage('Phân tích hoàn tất');
    } catch (err: any) {
      setError(`Lỗi PDF: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const matchResults = useMemo((): MatchResult[] => {
    const results: MatchResult[] = [];
    const pdfMap = new Map<string, number[]>();
    
    pdfPages.forEach((page) => {
      if (page.orderId) {
        const existing = pdfMap.get(page.orderId) || [];
        existing.push(page.pageIndex);
        pdfMap.set(page.orderId, existing);
      }
    });

    excelOrders.forEach((excelOrder, idx) => {
      const pageIndices = pdfMap.get(excelOrder.orderId);
      
      if (!pageIndices || pageIndices.length === 0) {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: null,
          status: 'missing_in_pdf'
        });
      } else if (pageIndices.length > 1) {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: pageIndices[0],
          status: 'duplicate_in_pdf'
        });
      } else {
        results.push({
          excelIndex: idx,
          orderId: excelOrder.orderId,
          pdfPageIndex: pageIndices[0],
          status: 'matched'
        });
      }
    });

    return results;
  }, [excelOrders, pdfPages]);

  const stats = useMemo((): ProcessingStats => {
    const matchedCount = matchResults.filter(r => r.status === 'matched').length;
    const missingInPdfCount = matchResults.filter(r => r.status === 'missing_in_pdf').length;
    const excelIdSet = new Set(excelOrders.map(o => o.orderId));
    const extraInPdfCount = pdfPages.filter(p => p.orderId && !excelIdSet.has(p.orderId)).length;
    const errorPdfCount = pdfPages.filter(p => !p.orderId).length;

    return {
      totalExcelRows: excelOrders.length,
      uniqueOrders: excelOrders.length,
      totalPdfPages: pdfPages.length,
      matchedCount,
      missingInPdfCount,
      extraInPdfCount,
      errorPdfCount
    };
  }, [excelOrders, pdfPages, matchResults]);

  const generateSortedPdf = async () => {
    if (!pdfFile || matchResults.length === 0) return;
    setIsProcessing(true);
    setProgress(0);
    setStatusMessage('Đang khởi tạo file PDF mới...');

    try {
      const existingPdfBytes = await pdfFile.arrayBuffer();
      const originalPdfDoc = await PDFDocument.load(new Uint8Array(existingPdfBytes), { 
        ignoreEncryption: true 
      });
      const newPdfDoc = await PDFDocument.create();

      const exportItems = matchResults.filter(r => (r.status === 'matched' || r.status === 'duplicate_in_pdf') && r.pdfPageIndex !== null);

      for (let i = 0; i < exportItems.length; i++) {
        const result = exportItems[i];
        if (result.pdfPageIndex !== null) {
          const [copiedPage] = await newPdfDoc.copyPages(originalPdfDoc, [result.pdfPageIndex]);
          newPdfDoc.addPage(copiedPage);
        }
        setStatusMessage(`Đang thêm trang ${i + 1}/${exportItems.length}...`);
        setProgress(Math.round(((i + 1) / exportItems.length) * 100));
      }

      setStatusMessage('Gói file dữ liệu...');
      const pdfBytes = await newPdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const originalName = pdfFile.name.replace('.pdf', '');
      link.href = url;
      link.download = `${originalName}_sorted.pdf`;
      link.click();
      setStatusMessage('Hoàn tất. File đã được tải về.');
    } catch (err: any) {
      setError(`Lỗi tạo PDF: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="bg-slate-50 w-full min-h-screen flex flex-col font-sans overflow-x-hidden text-slate-900 selection:bg-indigo-100">
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row justify-between items-center shrink-0 shadow-sm sticky top-0 z-50">
        <div className="flex items-center mb-4 sm:mb-0">
          <div className="size-[50px] bg-slate-100 rounded-2xl flex items-center justify-center shadow-inner mr-4 shrink-0 overflow-hidden border border-slate-200">
            <img 
              src="https://i.postimg.cc/7ZLc3GnZ/anh-logot.png" 
              alt="New Era Food Logo" 
              className="h-full w-full object-contain scale-[1.2]" 
              referrerPolicy="no-referrer" 
            />
          </div>
          <div>
            <h1 className="text-xl font-black tracking-tight text-slate-800 italic leading-tight">XỬ LÝ ĐƠN NỘI BỘ</h1>
            <p className="text-[10px] text-slate-400 uppercase font-black tracking-widest leading-none">v2.6 • High Density Mode</p>
          </div>
        </div>
        <div className="flex items-center space-x-4">
        </div>
      </header>

      {/* Tab Navigation */}
      <nav className="bg-slate-50/50 px-6 pt-2 flex items-center space-x-2 shrink-0 z-40">
        <button 
          onClick={() => { setActiveTab('pdf'); setStep(1); }}
          className={cn(
            "px-8 py-3 text-[11px] font-black uppercase tracking-widest transition-all duration-300 rounded-t-xl border-t border-x",
            activeTab === 'pdf' 
              ? "bg-white border-slate-200 text-indigo-600 shadow-[0_-4px_12px_-4px_rgba(0,0,0,0.05)]" 
              : "bg-transparent border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-100"
          )}
        >
          Sắp xếp PDF
        </button>
        <button 
          onClick={() => { setActiveTab('excel'); setStep(1); }}
          className={cn(
            "px-8 py-3 text-[11px] font-black uppercase tracking-widest transition-all duration-300 rounded-t-xl border-t border-x",
            activeTab === 'excel' 
              ? "bg-white border-slate-200 text-indigo-600 shadow-[0_-4px_12px_-4px_rgba(0,0,0,0.05)]" 
              : "bg-transparent border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-100"
          )}
        >
          Lọc Excel
        </button>
      </nav>

      <main className="flex-1 flex flex-col p-6 space-y-6 max-w-7xl mx-auto w-full overflow-hidden bg-white shadow-[0_0_50px_-12px_rgba(0,0,0,0.05)] border-x border-slate-100">
        {activeTab === 'pdf' ? (
          <>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 h-auto shrink-0">
              <motion.div 
                whileHover={{ y: -4, boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.05), 0 8px 10px -6px rgb(0 0 0 / 0.05)" }}
                className={cn(
                  "group bg-white border-2 border-dashed rounded-3xl p-8 flex flex-col items-center justify-center space-y-4 transition-all duration-300 relative",
                  excelFile ? "border-emerald-600 bg-emerald-50" : "border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/5"
                )}
              >
                <input type="file" accept=".xlsx, .xls, .csv" onChange={handleExcelUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={excelInputRef} />
                <div className={cn("p-5 rounded-2xl transition-all duration-500 shadow-sm", excelFile ? "bg-emerald-100 text-emerald-700 scale-110 rotate-3 shadow-emerald-100" : "bg-slate-50 text-slate-400 group-hover:scale-110 group-hover:rotate-3 shadow-slate-100")}>
                  <FileSpreadsheet className="w-10 h-10" />
                </div>
                <div className="text-center">
                  <p className="text-sm font-black text-slate-800 uppercase tracking-wide">1. TẢI LÊN EXCEL / CSV</p>
                  <p className="text-[11px] text-slate-400 font-bold max-w-[240px] mt-2 leading-relaxed">
                    {excelFile ? <span className="text-emerald-700 italic font-black truncate block">{excelFile.name}</span> : "Kéo thả file .xlsx, .csv (Tối đa 100MB)"}
                  </p>
                </div>
                {excelOrders.length > 0 && (
                  <motion.span 
                    initial={{ scale: 0.8 }} animate={{ scale: 1 }}
                    className="text-[10px] bg-emerald-500 text-white px-4 py-1.5 rounded-full font-black uppercase tracking-widest shadow-lg shadow-emerald-200"
                  >
                    ĐÃ NHẬN {excelOrders.length} Dòng
                  </motion.span>
                )}
              </motion.div>

              <motion.div 
                whileHover={{ y: -4, boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.05), 0 8px 10px -6px rgb(0 0 0 / 0.05)" }}
                className={cn(
                  "group bg-white border-2 border-dashed rounded-3xl p-8 flex flex-col items-center justify-center space-y-4 transition-all duration-300 relative",
                  pdfFile ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/5"
                )}
              >
                <input type="file" accept=".pdf" onChange={handlePdfUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={pdfInputRef} />
                <div className={cn("p-5 rounded-2xl transition-all duration-500 shadow-sm", pdfFile ? "bg-rose-100 text-rose-600 scale-110 -rotate-3 shadow-rose-100" : "bg-slate-50 text-slate-400 group-hover:scale-110 group-hover:-rotate-3 shadow-slate-100")}>
                  <FileText className="w-10 h-10" />
                </div>
                <div className="text-center">
                  <p className="text-sm font-black text-slate-800 uppercase tracking-wide">2. TẢI LÊN PDF GỐC</p>
                  <p className="text-[11px] text-slate-400 font-bold max-w-[240px] mt-2 leading-relaxed">
                    {pdfFile ? <span className="text-rose-600 italic font-black truncate block">{pdfFile.name}</span> : "Kéo thả file .pdf (Tối đa 100MB)"}
                  </p>
                </div>
                {pdfPages.length > 0 && (
                  <motion.span 
                    initial={{ scale: 0.8 }} animate={{ scale: 1 }}
                    className="text-[10px] bg-indigo-500 text-white px-4 py-1.5 rounded-full font-black uppercase tracking-widest shadow-lg shadow-indigo-200"
                  >
                    ĐÃ NHẬN {pdfPages.length} TRANG GỐC
                  </motion.span>
                )}
              </motion.div>
            </div>

            {(excelOrders.length > 0 && pdfFile && step === 1) && (
              <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} className="flex justify-center py-4">
                <motion.button 
                  whileHover={{ scale: 1.02, y: -2, boxShadow: "0 25px 50px -12px rgba(99, 102, 241, 0.25)" }}
                  whileTap={{ scale: 0.98 }}
                  onClick={processFiles}
                  disabled={isProcessing}
                  className="bg-indigo-600 text-white px-10 py-5 rounded-2xl font-black text-sm transition-all shadow-xl flex items-center gap-4 uppercase tracking-widest hover:bg-indigo-700 disabled:opacity-50 disabled:grayscale"
                >
                  {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Search className="w-5 h-5" />}
                  BẮT ĐẦU SẮP XẾP FILE PDF
                </motion.button>
              </motion.div>
            )}

            <div className="bg-slate-50/50 rounded-2xl p-6 shrink-0 border border-slate-100 space-y-4">
              <div className="flex flex-col lg:flex-row justify-between items-stretch lg:items-center gap-6">
                <div className="flex-1 space-y-2">
                  <div className="flex justify-between items-end">
                    <div className="flex items-center gap-3">
                      <div className="p-2 bg-indigo-100 rounded-lg">
                        {isProcessing ? <Loader2 className="w-4 h-4 animate-spin text-indigo-600" /> : <div className="w-4 h-4 rounded-full bg-indigo-600" />}
                      </div>
                      <p className="text-[11px] font-black text-slate-400 uppercase tracking-[0.1em]">
                        {statusMessage || (step === 2 ? "Hoàn tất đối chiếu" : "Chờ bắt đầu")}
                      </p>
                    </div>
                    <span className="text-xl font-black text-indigo-600 font-mono tracking-tighter">{progress}%</span>
                  </div>
                  <div className="w-full bg-white h-3.5 rounded-full overflow-hidden border border-slate-200 shadow-sm p-0.5">
                    <motion.div 
                      className="bg-gradient-to-r from-indigo-500 to-indigo-700 h-full rounded-full shadow-[0_0_10px_rgba(99,102,241,0.5)]"
                      initial={{ width: 0 }}
                      animate={{ width: `${progress}%` }}
                      transition={{ duration: 0.5, ease: "easeOut" }}
                    />
                  </div>
                </div>
                
                {step === 2 && (
                  <div className="grid grid-cols-2 lg:flex items-center gap-4 lg:gap-8 lg:border-l lg:border-slate-200 lg:pl-8">
                    <div className="bg-white p-3 rounded-2xl border border-slate-100 shadow-sm min-w-[100px]">
                      <p className="text-[9px] uppercase font-black text-slate-400 tracking-wider mb-1">Số dòng Excel</p>
                      <p className="text-xl font-black text-slate-800 tracking-tighter">{stats.uniqueOrders}</p>
                    </div>
                    <div className="bg-emerald-50 p-3 rounded-2xl border border-emerald-100 shadow-sm min-w-[100px]">
                      <p className="text-[9px] uppercase font-black text-emerald-600 tracking-wider mb-1">Đã khớp</p>
                      <p className="text-xl font-black text-emerald-600 tracking-tighter">{stats.matchedCount}</p>
                    </div>
                    <div className="bg-rose-50 p-3 rounded-2xl border border-rose-100 shadow-sm min-w-[100px]">
                      <p className="text-[9px] uppercase font-black text-rose-500 tracking-wider mb-1">Thiếu PDF</p>
                      <p className="text-xl font-black text-rose-500 tracking-tighter">{stats.missingInPdfCount}</p>
                    </div>
                    <div className="bg-amber-50 p-3 rounded-2xl border border-amber-100 shadow-sm min-w-[100px]">
                      <p className="text-[9px] uppercase font-black text-amber-500 tracking-wider mb-1">Lỗi / Dư</p>
                      <p className="text-xl font-black text-amber-500 tracking-tighter">{stats.errorPdfCount + stats.extraInPdfCount}</p>
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className="flex-1 bg-white rounded-3xl border border-slate-200 overflow-hidden flex flex-col min-h-[400px] shadow-sm">
              <div className="bg-slate-50 border-b border-slate-200 px-6 py-4 flex justify-between items-center shrink-0">
                <div className="flex items-center gap-3">
                  <div className="bg-indigo-600 p-1.5 rounded-lg">
                    <TableIcon className="w-4 h-4 text-white" />
                  </div>
                  <h3 className="text-[11px] font-black text-slate-700 uppercase tracking-[0.1em]">
                    Bảng Đối Chiếu Dữ Liệu
                  </h3>
                </div>
              </div>
              <div className="flex-1 overflow-auto bg-[radial-gradient(#f1f5f9_1px,transparent_1px)] [background-size:24px_24px]">
                {step === 1 ? (
                  <div className="h-full flex flex-col items-center justify-center p-12 text-center space-y-6">
                    <div className="w-20 h-20 bg-slate-50 rounded-3xl flex items-center justify-center border border-slate-100">
                      <Search className="w-10 h-10 text-slate-200" />
                    </div>
                    <div className="space-y-2">
                      <p className="text-xs font-black text-slate-400 uppercase tracking-widest">Đang chờ cấu trúc file...</p>
                      <p className="text-[10px] text-slate-300 font-bold max-w-[200px] mx-auto uppercase">Vui lòng tải lên cả file Excel và PDF để bắt đầu</p>
                    </div>
                  </div>
                ) : (
                  <table className="w-full text-left border-collapse table-fixed">
                    <thead className="sticky top-0 bg-white shadow-[0_2px_10px_-2px_rgba(0,0,0,0.05)] z-20">
                      <tr className="bg-slate-50/50 backdrop-blur-md">
                        <th className="p-4 text-[10px] font-black text-slate-500 uppercase w-20 text-center border-b border-slate-200">#</th>
                        <th className="p-4 text-[10px] font-black text-slate-500 uppercase border-b border-slate-200">Mã Đơn Hàng (Data Node)</th>
                        <th className="p-4 text-[10px] font-black text-slate-500 uppercase w-32 text-center border-b border-slate-200">Trang PDF</th>
                        <th className="p-4 text-[10px] font-black text-slate-500 uppercase border-b border-slate-200">Trạng Thái Kết Nối</th>
                      </tr>
                    </thead>
                    <tbody className="text-xs">
                      {matchResults.map((result, i) => (
                        <tr key={i} className={cn(
                          "transition-all duration-200 group",
                          i % 2 === 0 ? "bg-white" : "bg-slate-50/30",
                          "hover:bg-indigo-50/30"
                        )}>
                          <td className="p-4 text-center font-mono font-bold text-slate-400 group-hover:text-indigo-400">
                            {String(i + 1).padStart(2, '0')}
                          </td>
                          <td className="p-4 font-mono font-bold text-slate-700 group-hover:translate-x-1 transition-transform">
                            {result.orderId}
                          </td>
                          <td className="p-4 text-center">
                            {result.pdfPageIndex !== null ? (
                              <span className="bg-slate-100 text-slate-700 px-3 py-1 rounded-lg font-mono font-bold border border-slate-200">
                                {String(result.pdfPageIndex + 1).padStart(2, '0')}
                              </span>
                            ) : (
                              <span className="text-slate-300 font-mono font-bold">--</span>
                            )}
                          </td>
                          <td className="p-4">
                            {result.status === 'matched' ? (
                              <div className="inline-flex items-center gap-2">
                                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div>
                                <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest">Đã khớp</span>
                              </div>
                            ) : result.status === 'duplicate_in_pdf' ? (
                              <div className="inline-flex items-center gap-2">
                                <div className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse"></div>
                                <span className="text-[10px] font-black text-amber-600 uppercase tracking-widest">Data Trùng</span>
                              </div>
                            ) : (
                              <div className="inline-flex items-center gap-2">
                                <div className="w-1.5 h-1.5 rounded-full bg-rose-500"></div>
                                <span className="text-[10px] font-black text-rose-600 uppercase tracking-widest">Không tìm thấy</span>
                              </div>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>

            <div className="flex flex-col sm:flex-row justify-between items-center bg-slate-50 p-6 rounded-3xl border border-slate-100 gap-6 mt-auto">
              <div className="flex gap-4 w-full sm:w-auto">
                <motion.button 
                  whileHover={{ scale: 1.02, backgroundColor: "#f8fafc" }} 
                  whileTap={{ scale: 0.98 }}
                  onClick={resetAll} 
                  className="flex-1 sm:flex-none flex items-center justify-center gap-3 bg-white text-slate-500 px-8 py-4 rounded-2xl font-black text-[11px] border border-slate-200 transition-all uppercase tracking-widest shadow-sm hover:text-slate-700"
                >
                  <RefreshCcw className="w-4 h-4" /> Làm lại
                </motion.button>
              </div>
              
              <div className="flex items-center gap-6 w-full sm:w-auto">
                {step === 2 && (
                  <>
                    <div className="text-right hidden md:block">
                      <p className={cn("text-[10px] font-black uppercase tracking-widest mb-1", stats.missingInPdfCount > 0 ? "text-rose-500" : "text-emerald-500")}>
                        {stats.missingInPdfCount > 0 ? `${stats.missingInPdfCount} ĐƠN KHÔNG TÌM THẤY` : "ĐỒNG BỘ DỮ LIỆU HOÀN TẤT"}
                      </p>
                      <p className="text-[9px] text-slate-400 uppercase font-bold tracking-tight">Sắp xếp PDF tự động theo trình tự Excel</p>
                    </div>
                    <motion.button 
                      whileHover={{ scale: 1.02, y: -2, boxShadow: "0 25px 50px -12px rgba(79, 70, 229, 0.25)" }}
                      whileTap={{ scale: 0.98 }}
                      onClick={generateSortedPdf}
                      disabled={stats.matchedCount === 0 || isProcessing}
                      className="flex-1 sm:flex-none bg-indigo-600 text-white px-12 py-4 rounded-2xl font-black text-xs shadow-xl shadow-indigo-100 flex items-center justify-center gap-4 uppercase tracking-widest hover:bg-indigo-700 transition-all disabled:opacity-30"
                    >
                      {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5" />}
                      XUẤT PDF
                    </motion.button>
                  </>
                )}
              </div>
            </div>
          </>
        ) : (
          <div className="flex-1 flex flex-col space-y-6">
             {/* Excel Refiner Tool UI */}
             <motion.div 
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                className="bg-white rounded-[2rem] border border-slate-200 p-12 flex flex-col items-center justify-center text-center space-y-10 shadow-sm"
              >
                <div className={cn(
                  "group relative border-2 border-dashed rounded-[2.5rem] p-16 w-full max-w-2xl transition-all duration-500 shadow-inner",
                  rawExcelFile ? "border-emerald-600 bg-emerald-50" : "border-slate-100 hover:border-indigo-200 bg-slate-50/30"
                )}>
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleRawExcelUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" ref={excelRefinerInputRef} />
                  <div className={cn("size-24 rounded-3xl mx-auto mb-8 flex items-center justify-center shadow-lg transition-all duration-500", rawExcelFile ? "bg-emerald-600 text-white shadow-emerald-200 scale-110 rotate-6" : "bg-white text-slate-200 group-hover:scale-110")}>
                    <FileSpreadsheet className="w-12 h-12" />
                  </div>
                  <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tight">
                    TẢI LÊN FILE EXCEL THÔ
                  </h2>
                  <p className="text-sm text-slate-400 mt-3 font-medium max-w-sm mx-auto leading-relaxed">
                    Kéo thả file .xlsx, .csv (Tối đa 100MB)
                  </p>
                  
                  {rawExcelFile && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                      className="mt-8 px-6 py-3 bg-white rounded-2xl shadow-sm border border-indigo-100 inline-flex items-center gap-3"
                    >
                      <div className="size-2 rounded-full bg-emerald-500 animate-pulse"></div>
                      <span className="text-[11px] font-black text-slate-700 uppercase tracking-widest truncate max-w-[200px]">{rawExcelFile.name}</span>
                    </motion.div>
                  )}
                </div>

                {rawExcelData.length > 0 && (
                  <motion.div 
                    initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }}
                    className="w-full max-w-md bg-emerald-50/50 p-8 rounded-[2rem] border border-emerald-100 flex items-center gap-6 text-left"
                  >
                    <div className="bg-emerald-100 p-4 rounded-2xl shrink-0">
                      <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                    </div>
                    <div>
                      <h3 className="text-[11px] font-black uppercase text-emerald-700 tracking-widest mb-1">XỬ LÝ FILE EXCEL THÀNH CÔNG</h3>
                      <p className="text-lg font-black text-emerald-950 tracking-tighter leading-tight">{rawExcelData.length} Bản ghi dữ liệu</p>
                      <p className="text-[10px] text-emerald-600 font-bold uppercase mt-1">Sẵn sàng xuất file Standard</p>
                    </div>
                  </motion.div>
                )}

                <div className="flex flex-wrap justify-center gap-4">
                  {rawExcelFile && (
                    <motion.button 
                      whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                      onClick={resetAll} 
                      className="px-10 py-5 rounded-2xl bg-slate-50 text-slate-400 font-black text-[11px] uppercase tracking-widest border border-slate-100 hover:bg-white hover:text-slate-600 transition-all"
                    >
                      Làm lại
                    </motion.button>
                  )}
                  <motion.button 
                    whileHover={{ scale: 1.02, y: -4, boxShadow: "0 25px 50px -12px rgba(79, 70, 229, 0.25)" }}
                    whileTap={{ scale: 0.98 }}
                    onClick={exportCleanedExcel}
                    disabled={!rawExcelFile || isProcessing}
                    className="px-12 py-5 rounded-2xl bg-indigo-600 text-white font-black text-sm uppercase tracking-widest shadow-xl shadow-indigo-50 hover:bg-indigo-700 disabled:opacity-30 disabled:grayscale transition-all flex items-center gap-4"
                  >
                    {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5" />}
                    Tải file chuẩn
                  </motion.button>
                </div>
              </motion.div>
          </div>
        )}

        <AnimatePresence>
          {error && (
            <motion.div initial={{ y: 50, opacity: 0 }} animate={{ y: 0, opacity: 1 }} exit={{ y: 50, opacity: 0 }} className="fixed bottom-6 left-6 right-6 z-50 p-4 bg-rose-600 text-white rounded-2xl shadow-xl flex items-center justify-between border border-white/20">
              <div className="flex items-center gap-3">
                <AlertTriangle className="w-6 h-6 flex-shrink-0" />
                <p className="text-xs font-bold">{error}</p>
              </div>
              <button onClick={() => setError(null)} className="p-1 hover:bg-white/10 rounded-full">×</button>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
