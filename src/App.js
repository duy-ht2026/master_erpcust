import React, { useState, useEffect } from 'react';
import { 
  Upload, Download, FileCheck, Database, RefreshCw, AlertCircle, ChevronRight, ExternalLink, X, AlertTriangle
} from 'lucide-react';

/**
 * PROJECT: Checker [ERP Customer]
 * VERSION: 2.3.2
 * UPDATES: 
 * - Fixed row.rowNumber background to bg-indigo-100
 * - isNew Customer CUST_CODE background to bg-yellow-100 + font-bold
 * - Strengthened row colors: bg-red-100, bg-slate-100, hover:bg-emerald-100
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
    const cCode = row[1]?.toString().trim() || "";
    const cName = row[2]?.toString().trim() || "";
    const pXa = row[5]?.toString().trim() || "";
    const tCode = row[12]?.toString().trim() || "";

    // Dừng xử lý nếu 4 cột chính đều trống
    if (!cCode && !cName && !pXa && !tCode) {
      return "STOP_PROCESSING"; 
    }

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
          const results = [];
          for (let i = 0; i < rowsToProcess.length; i++) {
            const res = await processRow(rowsToProcess[i], i);
            if (res === "STOP_PROCESSING") break; 
            if (res !== null) results.push(res);
          }
          setData(results);
        }
      } catch (err) { 
        console.error("Lỗi xử lý file:", err); 
      } finally { 
        setIsProcessing(false); 
      }
    };
    reader.readAsBinaryString(file);
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
          <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"></div>
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-md overflow-hidden relative z-10 border border-slate-200">
            <div className={`p-5 flex flex-col items-center gap-4 text-center ${errorCount > 0 ? 'bg-red-50' : 'bg-emerald-50'}`}>
               <div className={`p-4 rounded-full ${errorCount > 0 ? 'bg-red-100 text-red-600' : 'bg-emerald-100 text-emerald-600'}`}>
                 {errorCount > 0 ? <AlertTriangle size={32} /> : <Download size={32} />}
               </div>
               <div>
                  <h3 className={`text-lg font-black uppercase ${errorCount > 0 ? 'text-red-700' : 'text-emerald-700'}`}>Xác nhận xuất dữ liệu</h3>
                  <p className="text-[10px] text-slate-500 font-bold mt-1">Hệ thống sẽ tạo file Excel theo template ERP</p>
               </div>
            </div>
            <div className="p-6">
              {errorCount > 0 && (
                <div className="bg-red-100/80 border border-red-200 rounded-lg p-3 mb-4 text-[11px] leading-relaxed text-red-800">
                  <span className="font-black uppercase block mb-1 underline">Cảnh báo lỗi!</span>
                  Dữ liệu chứa {errorCount} hàng chưa chuẩn. Bạn có chắc chắn muốn xuất không?
                </div>
              )}
              <div className="grid grid-cols-2 gap-3 mt-4">
                <button onClick={() => setShowConfirmModal(false)} className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-[10px] font-black uppercase transition-colors">Hủy</button>
                <button onClick={executeExport} className={`px-4 py-2 text-white rounded-lg text-[10px] font-black uppercase ${errorCount > 0 ? 'bg-red-600 hover:bg-red-700' : 'bg-[#057a10] hover:bg-emerald-700'}`}>Đồng ý</button>
              </div>
            </div>
          </div>
        </div>
      )}

      <header className="h-11 bg-[#057a10] text-white px-4 flex items-center justify-between shrink-0 shadow-md z-50">
        <div className="flex items-center gap-2.5">
          <FileCheck size={18} />
          <h1 className="text-[12px] font-black uppercase tracking-tight">Checker [ERP Customer]</h1>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-2 py-1 rounded-md bg-black/10 text-[9px] font-bold">
            <div className={`w-1.5 h-1.5 rounded-full ${dbStatus === 'online' ? 'bg-emerald-400' : 'bg-red-500'}`}></div>
            <span>{dbStatus === 'online' ? 'DB Active' : 'Connecting...'}</span>
            <RefreshCw size={10} className="cursor-pointer" onClick={fetchMasterData}/>
          </div>
          <input type="file" id="xlUp" className="hidden" onChange={handleFileUpload} accept=".xlsx, .xls" />
          <label htmlFor="xlUp" className="cursor-pointer px-3 py-1.5 bg-white/10 hover:bg-white/20 rounded text-[9px] font-black uppercase border border-white/20">Nạp Excel</label>
          <button onClick={() => setShowConfirmModal(true)} disabled={data.length === 0} className="px-3 py-1.5 bg-white text-[#057a10] rounded text-[9px] font-black uppercase disabled:opacity-30">Xuất File</button>
          <a href="https://checkmst.vercel.app/" target="_blank" className="text-white hover:underline text-[9px] font-bold uppercase flex items-center gap-1">Kiểm MST <ExternalLink size={10}/></a>
        </div>
      </header>

      <main className="flex-1 overflow-hidden p-2">
        {data.length > 0 ? (
          <div className="bg-white rounded-lg border border-slate-300 shadow-sm h-full flex flex-col overflow-hidden">
            <div className="px-3 py-2 bg-slate-50 border-b border-slate-300 flex items-center justify-between text-[10px] font-bold uppercase text-slate-500 shrink-0">
               <div>File: {fileName}</div>
               <div className="flex gap-4">
                 <span>Tổng: {data.length}</span>
                 <span className={errorCount > 0 ? "text-red-600" : "text-emerald-600"}>{errorCount} Lỗi</span>
               </div>
            </div>

            <div className="flex-1 overflow-auto">
              <table className="w-full border-separate border-spacing-0 text-[10.5px]">
                <thead>
                  <tr className="bg-slate-800 text-slate-300 sticky top-0 z-40">
                    <th className="w-10 px-1 py-1.5 text-center border-b border-r border-slate-600 font-bold uppercase sticky left-0 z-50 bg-slate-900">#</th>
                    {COLUMN_NAMES.map((h, i) => (
                      <th key={i} className="px-2 py-1.5 text-left border-b border-r border-slate-600 font-bold uppercase whitespace-nowrap min-w-[100px]">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-300">
                  {data.map((row, idx) => (
                    <tr key={idx} className={`
                      ${row.isError ? 'bg-red-100' : (idx % 2 === 1 ? 'bg-slate-100' : 'bg-white')}
                      hover:bg-emerald-100 transition-colors h-[28px]
                    `}>
                      <td className="px-1 py-1 text-center font-bold border-r border-slate-300 sticky left-0 z-20 bg-indigo-100 text-indigo-700">
                        {row.rowNumber}
                      </td>
                      {row.content.map((cell, cIdx) => {
                        const isValErr = row.validateErrors.includes(cIdx);
                        const isDbErr = row.databaseErrors.includes(cIdx);
                        const isMap = row.mappedIndices.includes(cIdx);
                        const isNew = row.isNewCustomer && cIdx === 1;
                        
                        let tdClass = `px-2 py-1 border-r border-slate-300 whitespace-nowrap overflow-hidden text-ellipsis max-w-[250px]`;
                        
                        if (isValErr) tdClass += " bg-red-600 text-white font-bold";
                        else if (isDbErr) tdClass += " bg-amber-500 text-white font-bold";
                        else if (isMap) tdClass += " bg-indigo-50 text-indigo-900 font-bold";
                        else if (isNew) tdClass += " bg-yellow-100 font-bold text-black";

                        return (
                          <td key={cIdx} className={tdClass} title={cell}>
                            {isValErr ? (cell || "TRỐNG") : isDbErr ? (cell || "KO_KHỚP") : cell}
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
          <div className="h-full flex flex-col items-center justify-center bg-white rounded-xl border-2 border-dashed border-slate-300">
            <Database size={48} className="text-slate-200 mb-3" />
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Chưa có dữ liệu. Hãy nạp file Excel.</span>
          </div>
        )}
      </main>

      <footer className="h-6 bg-white border-t border-slate-300 px-4 flex items-center justify-between text-[9px] text-slate-400 font-bold uppercase shrink-0">
        <div>© {new Date().getFullYear()} [IT-Master HTCorp]</div>
        <a href="https://masterdb.vercel.app/" target="_blank" className="flex items-center gap-1.5 hover:text-emerald-600">
          <div className={`w-1.5 h-1.5 rounded-full ${dbStatus === 'online' ? 'bg-emerald-500' : 'bg-red-500'}`}></div>
          <span>DB Synchronized</span>
          <ExternalLink size={10}/>
        </a>
      </footer>
    </div>
  );
};

const App = () => <Checker10 />;
export default App;
