import React, { useState, useEffect } from 'react';
import { 
  Upload, Download, FileCheck, Database, RefreshCw, AlertCircle, ChevronRight, ExternalLink, X, AlertTriangle
} from 'lucide-react';

/**
 * PROJECT: Checker [ERP Customer]
 * VERSION: 2.2.0
 * UPDATES: 
 * - Custom Modal for Export Confirmation with Error Warning
 * - Excel Export: Bold & Red text for Row 1 and Row 2
 * - Keep Bold text for specific columns in UI
 * - Keep DATE logic: Clear if not a valid date
 */

const SUPABASE_URL = "https://etdnpahmxdeurxlcuwcu.supabase.co";
const SUPABASE_KEY = "sb_publishable_vVs25rvLSgZXVkxw9WeT5w_xtaagYYG";

const COLUMN_NAMES = [
  "DATE", "CUST_CODE", "CUST_NAME", "CUST_NAME_FULL", "CUST_ADDRESS", 
  "PHUONG_XA_NEW", "PHUONG_XA_OLD", "CUST_PHONE", "CHANNEL", "BIZ_TYPE", 
  "CUST_GROUP", "SALEMAN", "TAX_CODE", "TAX_BUYER", "TAX_NAME", 
  "TAX_ADDRESS", "EMAIL", "PAYM_TERM", "COMMENT", "BIZ_CODE", "CUST_GRPCODE",
  "CODE_PHUONG", "CODE_CITY", "CODE_REGION"
];

const EXPORT_TEMPLATE_HEADERS = [
  "CUST_CODE", "CUST_NAME", "CUST_NAME_FULL", "TAX_NAME", "FIX_Loaikh", 
  "CUST_GRPCODE", "TAX_CODE", "FIX_Currency", "CHANNEL", "SALEMAN", 
  "FIX_BGcha", "BIZ_CODE", "CUST_ADDRESS", "PAYM_TERM", "PAYM_METHOD", 
  "TAX_ADDRESS", "FIX_Country", "CODE_CITY", "CODE_PHUONG", "CODE_REGION", 
  "TAX_BUYER", "CUST_PHONE", "FIX_ext", "FIX_fax", "FIX_Mobile", 
  "EMAIL", "FIX_Web", "Fix_Cmt", "FIX_A", "FIX_B", "FIX_CN"
];

const Checker10 = () => {
  const [data, setData] = useState([]);
  const [validChannels, setValidChannels] = useState([]);
  const [custGroups, setCustGroups] = useState([]);
  const [bizModels, setBizModels] = useState([]); 
  const [addressPreviewCount, setAddressPreviewCount] = useState(0); 
  const [fileName, setFileName] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [dbStatus, setDbStatus] = useState('connecting');
  const [supabaseClient, setSupabaseClient] = useState(null);
  const [showConfirmModal, setShowConfirmModal] = useState(false);

  useEffect(() => {
    if (!window.XLSX && !document.getElementById('xlsx-sdk')) {
      const xScript = document.createElement('script');
      xScript.id = 'xlsx-sdk';
      xScript.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      document.body.appendChild(xScript);
    }

    const initSupabase = () => {
      if (window.supabase) {
        const client = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
        setSupabaseClient(client);
      } else {
        const sScript = document.createElement('script');
        sScript.id = 'supabase-sdk';
        sScript.src = "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2";
        sScript.onload = () => {
          if (window.supabase) {
            const client = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
            setSupabaseClient(client);
          }
        };
        document.body.appendChild(sScript);
      }
    };
    initSupabase();
  }, []);

  useEffect(() => {
    if (supabaseClient) fetchMasterData();
  }, [supabaseClient]);

  const fetchMasterData = async () => {
    if (!supabaseClient) return;
    setDbStatus('loading');
    try {
      const { data: channels } = await supabaseClient.from('ConfigChannel').select('chcode');
      const { data: models } = await supabaseClient.from('ConfigModel').select('chanel_code, mdname, mdcode');
      const { data: groups } = await supabaseClient.from('ConfigCustGroup').select('chanel_code, group_name, group_code');
      const { count } = await supabaseClient.from('ConfigAddress').select('*', { count: 'exact', head: true });
      
      setValidChannels(channels?.map(c => c.chcode?.toString().toUpperCase() || "") || []);
      setBizModels(models || []);
      setCustGroups(groups || []);
      setAddressPreviewCount(count || 0);
      setDbStatus('online');
    } catch (err) {
      setDbStatus('error');
    }
  };

  const formatExcelDate = (val) => {
    if (!val) return "";
    if (typeof val === 'number' && val > 30000) {
      const date = new Date((val - 25569) * 86400 * 1000);
      if (!isNaN(date.getTime())) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
      }
    }
    if (typeof val === 'string' && val.trim() !== "") {
        const date = new Date(val);
        if (!isNaN(date.getTime())) {
            return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
        }
    }
    return "";
  };

  const findAddressOnline = async (addressText) => {
    if (!addressText || !supabaseClient) return null;
    try {
      const { data, error } = await supabaseClient
        .from('ConfigAddress')
        .select('code_phuong, code_city, code_region')
        .ilike('address_short', addressText.trim())
        .limit(1);
      
      if (error || !data || data.length === 0) return null;
      return data[0];
    } catch (e) {
      return null;
    }
  };

  const processRow = async (row, idx) => {
    if (row.length === 0 || row.every(cell => cell === "")) return null;

    const content = COLUMN_NAMES.map((_, cIdx) => {
      const cell = row[cIdx];
      if (cIdx === 0) return formatExcelDate(cell);
      let val = (cell === undefined || cell === null) ? "" : cell.toString().trim();
      if (cIdx === 1 && (val === "" || val.toUpperCase() === "NEW")) return "NEW";
      return val;
    });

    const validateErrors = new Set();
    const databaseErrors = new Set();
    const mappedIndices = new Set();

    const CUST_CODE = content[1]?.toString().toUpperCase();
    const PHUONG_XA_INPUT = content[5]?.toString().trim();
    const CHANNEL = content[8]?.toString().toUpperCase(); 
    const BIZ_TYPE = content[9]?.toString().toUpperCase();
    const GROUP_NAME_INPUT = content[10]?.toString().toUpperCase();

    if (!CHANNEL || CHANNEL === "") {
      validateErrors.add(8);
    } else if (!validChannels.includes(CHANNEL)) {
      databaseErrors.add(8);
    }

    if (CHANNEL && BIZ_TYPE) {
      const foundModel = bizModels.find(m => 
        m.chanel_code?.toString().toUpperCase() === CHANNEL && 
        m.mdname?.toString().toUpperCase() === BIZ_TYPE
      );
      if (foundModel) {
        content[19] = foundModel.mdcode; 
        mappedIndices.add(19);
      } else {
        databaseErrors.add(9);
        databaseErrors.add(19); 
        content[19] = "ERR_BIZ";
      }
    } else if (!BIZ_TYPE) validateErrors.add(9);

    if (CHANNEL && GROUP_NAME_INPUT) {
      const foundGroup = custGroups.find(g => 
        g.chanel_code?.toString().toUpperCase() === CHANNEL && 
        g.group_name?.toString().toUpperCase() === GROUP_NAME_INPUT
      );
      if (foundGroup) {
        content[20] = foundGroup.group_code;
        mappedIndices.add(20);
      } else {
        databaseErrors.add(10);
        databaseErrors.add(20);
        content[20] = "ERR_GRP";
      }
    } else if (!GROUP_NAME_INPUT) validateErrors.add(10);

    if (PHUONG_XA_INPUT) {
      const foundAddr = await findAddressOnline(PHUONG_XA_INPUT);
      if (foundAddr) {
        content[21] = foundAddr.code_phuong;
        content[22] = foundAddr.code_city;
        content[23] = foundAddr.code_region;
        mappedIndices.add(21);
        mappedIndices.add(22);
        mappedIndices.add(23);
      } else {
        databaseErrors.add(5);
        content[21] = "ERR_ADDR";
        content[22] = "ERR_ADDR";
        content[23] = "ERR_ADDR";
      }
    } else {
      validateErrors.add(5);
    }

    const isNewCustomer = CUST_CODE === "NEW";
    if (isNewCustomer) {
      const requiredForNew = [2, 3, 4, 5, 6, 8, 9, 10, 14, 15];
      requiredForNew.forEach(idx => {
        if (!content[idx] || content[idx].toString().trim() === "") {
          validateErrors.add(idx);
        }
      });
    }

    return { 
      content, 
      rowNumber: idx + 4, 
      isError: validateErrors.size > 0 || databaseErrors.size > 0, 
      isNewCustomer, 
      validateErrors: Array.from(validateErrors), 
      databaseErrors: Array.from(databaseErrors), 
      mappedIndices: Array.from(mappedIndices) 
    };
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file || !window.XLSX) return;
    setFileName(file.name);
    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = window.XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawJson = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        
        if (rawJson.length >= 3) {
          const rowsToProcess = rawJson.slice(3);
          const results = await Promise.all(
            rowsToProcess.map((row, idx) => processRow(row, idx))
          );
          setData(results.filter(item => item !== null));
        }
      } catch (err) { 
        console.error("Lỗi xử lý file:", err); 
      } finally { 
        setIsProcessing(false); 
      }
    };
    reader.readAsBinaryString(file);
  };

  const confirmExport = () => {
    if (data.length === 0) return;
    setShowConfirmModal(true);
  };

  const executeExport = () => {
    if (data.length === 0 || !window.XLSX) return;
    setShowConfirmModal(false);

    const exportRows = data.map(item => {
      const raw = item.content;
      return EXPORT_TEMPLATE_HEADERS.map(header => {
        switch (header) {
          case "CUST_CODE": return raw[1];
          case "CUST_NAME": return raw[2];
          case "CUST_NAME_FULL": return raw[3];
          case "TAX_NAME": return raw[14];
          case "FIX_Loaikh": return "1";
          case "CUST_GRPCODE": return raw[20];
          case "TAX_CODE": return raw[12];
          case "FIX_Currency": return "VND";
          case "CHANNEL": return raw[8];
          case "SALEMAN": return raw[11];
          case "FIX_BGcha": return "";
          case "BIZ_CODE": return raw[19];
          case "CUST_ADDRESS": return raw[4];
          case "PAYM_TERM": return "COD";
          case "PAYM_METHOD": return "TM/CK";
          case "TAX_ADDRESS": return raw[15];
          case "FIX_Country": return "VN";
          case "CODE_CITY": return raw[22];
          case "CODE_PHUONG": return raw[21];
          case "CODE_REGION": return raw[23];
          case "TAX_BUYER": return raw[13];
          case "CUST_PHONE": return raw[7];
          case "EMAIL": return raw[16];
          case "FIX_A": return "FALSE";
          case "FIX_B": return "0";
          case "FIX_CN": return "C";
          default: 
            if (header.startsWith("FIX_") || header.startsWith("Fix_")) return ""; 
            return "";
        }
      });
    });

    const finalAOA = [
      ["Dữ liệu này được export từ Web-App (Checker ERP Customer)"],
      EXPORT_TEMPLATE_HEADERS,
      ...exportRows
    ];

    const ws = window.XLSX.utils.aoa_to_sheet(finalAOA);

    // Bôi đậm và tô màu đỏ cho dòng 1 và dòng 2
    // Lưu ý: Tính năng Style yêu cầu thư viện xlsx-js-style hoặc SheetJS Pro.
    // Với thư viện cộng đồng 'xlsx', style thường bị bỏ qua, nhưng ta áp dụng cấu trúc cell object để hỗ trợ tốt nhất
    const range = window.XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      // Dòng 1 (Row 0)
      const cell1 = ws[window.XLSX.utils.encode_cell({ r: 0, c: C })];
      if (cell1) {
        cell1.s = {
          font: { bold: true, color: { rgb: "FF0000" } }
        };
      }
      // Dòng 2 (Row 1)
      const cell2 = ws[window.XLSX.utils.encode_cell({ r: 1, c: C })];
      if (cell2) {
        cell2.s = {
          font: { bold: true, color: { rgb: "FF0000" } }
        };
      }
    }

    const wscols = EXPORT_TEMPLATE_HEADERS.map(h => ({ wch: h.length + 10 }));
    ws['!cols'] = wscols;

    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "DMS_Customer_Template");
    window.XLSX.writeFile(wb, `erpCustomerExpWeb_${new Date().getTime()}.xlsx`);
  };

  const errorCount = data.filter(d => d.isError).length;

  return (
    <div className="flex flex-col h-screen bg-[#f8fafc] font-sans overflow-hidden text-slate-800">
      {/* Confirmation Modal */}
      {showConfirmModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200"></div>
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden relative z-10 border border-slate-200 animate-in zoom-in-95 duration-200">
            <div className={`p-5 flex flex-col items-center gap-4 text-center ${errorCount > 0 ? 'bg-red-50' : 'bg-emerald-50'}`}>
               <div className={`p-4 rounded-full ${errorCount > 0 ? 'bg-red-100 text-red-600' : 'bg-emerald-100 text-emerald-600'}`}>
                 {errorCount > 0 ? <AlertTriangle size={32} /> : <Download size={32} />}
               </div>
               <div>
                  <h3 className={`text-lg font-black uppercase ${errorCount > 0 ? 'text-red-700' : 'text-emerald-700'}`}>
                    Xác nhận xuất dữ liệu
                  </h3>
                  <p className="text-xs text-slate-500 font-bold mt-1">
                    Bạn đang chuẩn bị xuất file ERP Customer Template
                  </p>
               </div>
            </div>
            
            <div className="p-6">
              {errorCount > 0 ? (
                <div className="bg-red-100/50 border border-red-200 rounded-lg p-3 mb-4">
                  <div className="flex items-start gap-2">
                    <AlertCircle size={16} className="text-red-600 shrink-0 mt-0.5" />
                    <div className="text-[11px] leading-relaxed">
                      <span className="font-black text-red-700 uppercase underline block mb-1">Cảnh báo: Phát hiện {errorCount} dòng lỗi!</span>
                      Dữ liệu đang có các giá trị trống hoặc sai danh mục. File xuất ra có thể sẽ không được chấp nhận khi import vào hệ thống ERP. Bạn vẫn muốn tiếp tục?
                    </div>
                  </div>
                </div>
              ) : (
                <div className="bg-emerald-100/50 border border-emerald-200 rounded-lg p-3 mb-4">
                  <div className="flex items-center gap-2 text-emerald-700">
                    <FileCheck size={16} />
                    <span className="text-[11px] font-bold uppercase tracking-tight">Dữ liệu hiện tại đã sẵn sàng để xuất file.</span>
                  </div>
                </div>
              )}

              <div className="grid grid-cols-2 gap-3 mt-6">
                <button 
                  onClick={() => setShowConfirmModal(false)}
                  className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-[10px] font-black uppercase transition-colors"
                >
                  Hủy bỏ (Cancel)
                </button>
                <button 
                  onClick={executeExport}
                  className={`px-4 py-2.5 text-white rounded-lg text-[10px] font-black uppercase shadow-lg shadow-black/10 transition-all active:scale-95 ${errorCount > 0 ? 'bg-red-600 hover:bg-red-700' : 'bg-[#057a10] hover:bg-emerald-700'}`}
                >
                  Đồng ý (Ok)
                </button>
              </div>
            </div>
            <button 
              onClick={() => setShowConfirmModal(false)}
              className="absolute top-3 right-3 text-slate-400 hover:text-slate-600"
            >
              <X size={18} />
            </button>
          </div>
        </div>
      )}

      <header className="h-11 bg-[#057a10] text-white px-4 flex items-center justify-between shrink-0 shadow-md z-50">
        <div className="flex items-center gap-2.5">
          <div className="bg-white/20 p-1 rounded-md">
            <FileCheck size={18} className="text-white" />
          </div>
          <h1 className="text-[12px] font-black uppercase tracking-tight">Checker [ERP Customer]</h1>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-2.5 py-1 rounded-md bg-black/10 border border-white/10">
            <div className={`w-1.5 h-1.5 rounded-full ${dbStatus === 'online' ? 'bg-emerald-400' : 'bg-red-500'}`}></div>
            <span className="text-[9px] font-bold uppercase">{dbStatus === 'online' ? 'Database Active' : 'Connecting...'}</span>
            <button onClick={fetchMasterData} className="hover:rotate-180 transition-transform ml-0.5 opacity-70 hover:opacity-100">
              <RefreshCw size={10}/>
            </button>
          </div>

          <input type="file" id="xlUp" className="hidden" onChange={handleFileUpload} accept=".xlsx, .xls" />
          <label htmlFor="xlUp" className={`cursor-pointer px-4 py-1.5 bg-white/10 hover:bg-white/20 rounded text-[9px] font-black uppercase border border-white/20 transition-all ${isProcessing ? 'opacity-50 pointer-events-none' : ''}`}>
            {isProcessing ? 'Đang Xử Lý...' : 'Nạp Excel'}
          </label>
          
          <div className="flex items-center gap-2.5">
            <button 
              onClick={confirmExport} 
              disabled={data.length === 0 || isProcessing} 
              className="px-4 py-1.5 bg-white text-[#057a10] rounded text-[9px] font-black uppercase shadow-sm disabled:opacity-30 active:scale-95 transition-all"
            >
              Xuất File
            </button>
            <div className="w-[1px] h-4 bg-white/20 mx-1"></div>
            <a 
              href="https://checkmst.vercel.app/" 
              target="_blank" 
              rel="noopener noreferrer"
              className="flex items-center gap-1.5 text-white hover:text-emerald-100 transition-colors"
              title="Kiểm tra mã số thuế"
            >
              <ExternalLink size={14} />
              <span className="text-[9px] font-bold underline underline-offset-4 decoration-white/40 uppercase">Kiểm MST</span>
            </a>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-hidden p-2">
        {data.length > 0 ? (
          <div className="bg-white rounded-lg border border-slate-300 shadow-sm h-full flex flex-col overflow-hidden">
            <div className="px-3 py-2 bg-slate-50 border-b border-slate-300 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-4">
                 <div className="flex items-center gap-1.5 text-[10px] font-bold text-slate-700 uppercase">
                   <ChevronRight size={14} className="text-[#057a10]"/> {fileName}
                 </div>
                 <div className="flex items-center gap-4 border-l border-slate-300 pl-4">
                    <div className="flex items-center gap-1.5 text-[8px] font-bold text-slate-500 uppercase">
                        <div className="w-2 h-2 bg-red-600 rounded-sm"></div> Thiếu Thông Tin
                    </div>
                    <div className="flex items-center gap-1.5 text-[8px] font-bold text-slate-500 uppercase">
                        <div className="w-2 h-2 bg-amber-500 rounded-sm"></div> Sai Danh Mục
                    </div>
                    <div className="flex items-center gap-1.5 text-[8px] font-bold text-slate-500 uppercase">
                        <span className="text-red-600 font-black">NEW</span> = Khách Mới
                    </div>
                 </div>
              </div>
              <div className="flex items-center gap-3 shrink-0">
                 <span className="text-[9px] font-bold text-slate-400 uppercase">Hàng: {data.length}</span>
                 <span className={`px-2 py-0.5 rounded text-[9px] font-black uppercase ${errorCount > 0 ? 'bg-red-100 text-red-700' : 'bg-emerald-100 text-emerald-700'}`}>
                   {errorCount} Lỗi
                 </span>
              </div>
            </div>

            <div className="flex-1 overflow-auto relative">
              {isProcessing && (
                <div className="absolute inset-0 bg-white/60 z-50 flex items-center justify-center backdrop-blur-[1px]">
                    <div className="flex flex-col items-center gap-2">
                        <RefreshCw className="animate-spin text-[#057a10]" size={24} />
                        <span className="text-[9px] font-black uppercase tracking-widest text-[#057a10]">Đang đối soát địa chỉ Database...</span>
                    </div>
                </div>
              )}
              <table className="w-full border-separate border-spacing-0 text-[10.5px] min-w-max">
                <thead>
                  <tr className="bg-slate-800 text-slate-300 sticky top-0 z-40">
                    <th className="w-10 px-1 py-1.5 text-center border-b border-r border-slate-600 font-bold uppercase bg-slate-900 sticky left-0 z-50">#</th>
                    {COLUMN_NAMES.map((h, i) => (
                      <th 
                        key={i} 
                        style={{ width: i === 4 ? '280px' : 'auto', maxWidth: i === 4 ? '280px' : '250px' }}
                        className={`px-2 py-1.5 text-left border-b border-r border-slate-600 font-bold uppercase whitespace-nowrap
                          ${[5, 8, 9, 10].includes(i) ? 'bg-slate-700 text-emerald-300' : ''}
                          ${i >= 19 ? 'bg-indigo-950 text-indigo-200' : ''}
                        `}
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-300">
                  {data.map((row, idx) => (
                    <tr key={idx} className={`${row.isError ? 'bg-red-50/50' : 'bg-white'} hover:bg-slate-50 transition-colors h-[28px]`}>
                      <td className={`px-1 py-1 text-center font-bold border-r border-slate-300 sticky left-0 z-20 ${row.isError ? 'text-red-600 bg-red-50' : 'text-slate-400 bg-white'}`}>
                        {row.rowNumber}
                      </td>
                      {row.content.map((cell, cIdx) => {
                        const isValErr = row.validateErrors.includes(cIdx);
                        const isDbErr = row.databaseErrors.includes(cIdx);
                        const isMap = row.mappedIndices.includes(cIdx);
                        const isAddress = cIdx === 4;
                        const isErrAddr = cell === "ERR_ADDR";
                        
                        const isTargetBold = [1, 9, 10, 12, 14].includes(cIdx);

                        let tdClass = `px-2 py-1 border-r border-slate-300 align-middle`;
                        let style = isAddress ? { width: '280px', minWidth: '280px', maxWidth: '280px' } : { maxWidth: '250px' };
                        
                        if (isAddress) tdClass += " whitespace-normal leading-tight break-words";
                        else tdClass += " whitespace-nowrap overflow-hidden text-ellipsis text-[10px]";

                        if (isValErr) {
                          tdClass += " bg-red-600 text-white font-bold";
                        } else if (isDbErr) {
                          tdClass += " bg-amber-500 text-white font-bold";
                        } else if (isMap) {
                          tdClass += " bg-indigo-50 text-indigo-900 font-bold border-l border-indigo-300";
                        }

                        return (
                          <td key={cIdx} className={tdClass} style={style} title={cell}>
                            <span className={`
                              ${isTargetBold ? "font-bold text-slate-900" : ""}
                              ${cIdx === 1 && cell === "NEW" ? "text-red-600 font-black" : ""}
                              ${isErrAddr ? "text-amber-600 font-bold" : ""}
                            `}>
                              {isValErr ? (cell || "TRỐNG") : isDbErr ? (cell || "KO_KHỚP") : cell}
                            </span>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          <div className="h-full flex flex-col items-center justify-center bg-white rounded-xl border-2 border-dashed border-slate-300 shadow-sm">
            <div className="p-8 bg-emerald-50 rounded-full mb-5 border border-emerald-100 shadow-inner">
                <Database size={56} className="text-[#057a10] opacity-80" />
            </div>
            <h2 className="text-xl font-black uppercase text-slate-300 tracking-[0.2em] mb-4">Database Checker</h2>
            <div className="flex gap-3">
                <div className="px-5 py-2 bg-slate-50 rounded-xl border border-slate-200 flex flex-col items-center min-w-[100px]">
                    <span className="text-[8px] font-bold text-slate-400 uppercase">Channel Master</span>
                    <span className="text-base font-black text-slate-600">{validChannels.length}</span>
                </div>
                <div className="px-5 py-2 bg-slate-50 rounded-xl border border-slate-200 flex flex-col items-center min-w-[100px]">
                    <span className="text-[8px] font-bold text-slate-400 uppercase">Address Records</span>
                    <span className="text-base font-black text-slate-600">{addressPreviewCount.toLocaleString()}</span>
                </div>
                <div className="px-5 py-2 bg-slate-50 rounded-xl border border-slate-200 flex flex-col items-center min-w-[100px]">
                    <span className="text-[8px] font-bold text-slate-400 uppercase">Database Mode</span>
                    <span className="text-base font-black text-[#057a10]">CONNECTED</span>
                </div>
            </div>
          </div>
        )}
      </main>

      <footer className="h-6 bg-white border-t border-slate-300 px-4 flex items-center justify-between text-[9px] text-slate-400 font-bold uppercase tracking-wider shrink-0">
        <div>
          <span>© {new Date().getFullYear()} [IT-Master HTCorp] - ERP Customer Validation Tool</span>
        </div>
        <div className="flex items-center gap-2">
            <div className={`w-1.5 h-1.5 rounded-full ${dbStatus === 'online' ? 'bg-emerald-500' : 'bg-red-500'}`}></div>
            <span>{dbStatus === 'online' ? 'Database Synchronized' : 'Database Offline'}</span>
        </div>
      </footer>
    </div>
  );
};

const App = () => <Checker10 />;
export default App;
