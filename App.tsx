
import React, { useState, useMemo } from 'react';
import { Upload, FileSpreadsheet, Download, RefreshCw, CheckCircle2, AlertCircle, Calendar, Info, FileEdit, User } from 'lucide-react';
import { AttendanceRecord } from './types';
import { readExcel, processAttendanceData, exportToExcel } from './services/excelService';

const App: React.FC = () => {
  const [data, setData] = useState<AttendanceRecord[] | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setFileName(file.name);

    try {
      const result = await readExcel(file);
      const processed = processAttendanceData(result.data);
      setHeaders(result.headers);
      setData(processed);
    } catch (err) {
      console.error(err);
      setError('處理失敗：請檢查 Excel 格式與標題。');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = () => {
    if (!data || !fileName || headers.length === 0) return;
    exportToExcel(data, headers, fileName);
  };

  const reset = () => {
    setData(null);
    setHeaders([]);
    setFileName('');
    setError(null);
  };

  const summary = useMemo(() => {
    if (!data) return { lateCount: 0, earlyCount: 0, modifiedCount: 0 };
    return data.reduce((acc, row) => {
      if (row._modifiedFields?.has('遲到(分鐘)')) acc.lateCount++;
      if (row._modifiedFields?.has('早退(分鐘)')) acc.earlyCount++;
      if (row['異動'] === 'V') acc.modifiedCount++;
      return acc;
    }, { lateCount: 0, earlyCount: 0, modifiedCount: 0 });
  }, [data]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-20 font-sans">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-2 rounded-lg">
              <Calendar className="text-white w-5 h-5" />
            </div>
            <h1 className="text-lg font-bold tracking-tight text-slate-800">
              遠距出勤助手 <span className="text-slate-400 font-normal ml-2">週五邏輯處理</span>
            </h1>
          </div>
          {data && (
            <button 
              onClick={reset}
              className="text-sm font-medium text-slate-500 hover:text-indigo-600 flex items-center gap-1.5 py-2 px-3 hover:bg-slate-100 rounded-lg transition-all"
            >
              <RefreshCw className="w-4 h-4" />
              重新上傳
            </button>
          )}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8">
        {!data && !loading && (
          <div className="bg-white rounded-3xl border border-slate-200 p-12 text-center shadow-xl shadow-slate-200/50 animate-in fade-in zoom-in-95 duration-500">
            <div className="max-w-sm mx-auto">
              <div className="w-24 h-24 bg-indigo-50 rounded-3xl flex items-center justify-center mx-auto mb-8">
                <Upload className="w-12 h-12 text-indigo-600" />
              </div>
              <h2 className="text-2xl font-bold mb-3 text-slate-800">上傳考勤 Excel</h2>
              <p className="text-slate-500 mb-10 leading-relaxed">
                自動偵測週五記錄，依據「應出勤時數」動態計算遲到基準。<br/>
                <span className="text-indigo-600 font-bold">匯出將在最後一欄新增「異動 (V)」標記。</span>
              </p>
              
              <label className="group relative cursor-pointer inline-flex items-center justify-center px-10 py-5 bg-indigo-600 text-white font-bold rounded-2xl hover:bg-indigo-700 transition-all active:scale-95 shadow-xl shadow-indigo-200">
                <input 
                  type="file" 
                  accept=".xlsx, .xls" 
                  className="hidden" 
                  onChange={handleFileUpload}
                />
                選擇檔案並開始處理
              </label>
            </div>
          </div>
        )}

        {error && (
          <div className="bg-rose-50 border border-rose-200 text-rose-700 px-6 py-4 rounded-2xl flex items-center gap-3 mb-8">
            <AlertCircle className="w-5 h-5 flex-shrink-0" />
            <p className="text-sm font-semibold">{error}</p>
          </div>
        )}

        {loading && (
          <div className="flex flex-col items-center justify-center py-32 text-center">
            <div className="w-16 h-16 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin mb-6"></div>
            <p className="text-slate-500 font-bold tracking-wide">系統正在執行邏輯運算...</p>
          </div>
        )}

        {data && !loading && (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-6 duration-700">
            {/* Summary Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-amber-100 p-3 rounded-xl text-amber-700">
                  <Info className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs font-black text-slate-400 uppercase tracking-widest">週五遲到次數</p>
                  <p className="text-2xl font-black text-slate-800">{summary.lateCount}</p>
                </div>
              </div>
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-rose-100 p-3 rounded-xl text-rose-700">
                  <Info className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs font-black text-slate-400 uppercase tracking-widest">週五早退次數</p>
                  <p className="text-2xl font-black text-slate-800">{summary.earlyCount}</p>
                </div>
              </div>
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center gap-4">
                <div className="bg-indigo-100 p-3 rounded-xl text-indigo-700">
                  <FileEdit className="w-6 h-6" />
                </div>
                <div>
                  <p className="text-xs font-black text-slate-400 uppercase tracking-widest">總異動筆數 (V)</p>
                  <p className="text-2xl font-black text-slate-800">{summary.modifiedCount}</p>
                </div>
              </div>
            </div>

            {/* Action Bar */}
            <div className="bg-white rounded-3xl shadow-lg border border-slate-200 p-8 flex flex-col md:flex-row md:items-center justify-between gap-6">
              <div className="flex items-center gap-5">
                <div className="w-14 h-14 bg-emerald-50 rounded-2xl flex items-center justify-center border border-emerald-100">
                  <FileSpreadsheet className="text-emerald-600 w-7 h-7" />
                </div>
                <div>
                  <h3 className="font-bold text-lg text-slate-800">{fileName}</h3>
                  <p className="text-sm text-slate-500 mt-0.5">新增「異動」欄位並以 <span className="text-rose-600 font-bold">黃底紅字</span> 標示異動</p>
                </div>
              </div>
              
              <button 
                onClick={handleDownload}
                className="group inline-flex items-center justify-center gap-3 px-10 py-4 bg-slate-900 text-white font-bold rounded-2xl hover:bg-slate-800 transition-all shadow-lg active:scale-95"
              >
                <Download className="w-5 h-5" />
                下載處理後的 Excel
              </button>
            </div>

            {/* Preview Table */}
            <div className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="p-6 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                <h4 className="font-bold text-slate-700 flex items-center gap-2">
                  <CheckCircle2 className="w-5 h-5 text-indigo-600" />
                  數據預覽 (前 20 筆)
                </h4>
                <div className="flex gap-4 text-[11px] font-bold">
                  <span className="flex items-center gap-1.5">
                    <span className="w-3 h-3 bg-yellow-200 border border-yellow-300 rounded"></span> 異動標示
                  </span>
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-white border-b border-slate-200">
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase">員工編號</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase">日期</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase text-center">上班時間</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase text-center">下班時間</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase text-center">遲到(分)</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase text-center">早退(分)</th>
                      <th className="px-6 py-4 text-xs font-black text-slate-400 uppercase text-center">異動</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {data.slice(0, 20).map((row, idx) => {
                      const dateVal = row['出勤日期'];
                      let isFriday = false;
                      let displayDate = '-';
                      
                      try {
                        const d = (dateVal instanceof Date) ? dateVal : new Date(String(dateVal));
                        if (!isNaN(d.getTime())) {
                          isFriday = (d.getDay() === 5);
                          displayDate = `${d.getFullYear()}/${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getDate().toString().padStart(2, '0')}`;
                        }
                      } catch (e) {}
                      
                      const mods = row._modifiedFields || new Set();

                      return (
                        <tr key={idx} className={`${isFriday ? 'bg-indigo-50/20' : ''} hover:bg-slate-50 transition-colors`}>
                          <td className="px-6 py-4">
                            <div className="flex items-center gap-2">
                              <User className="w-3.5 h-3.5 text-slate-400" />
                              <span className="text-sm font-bold text-slate-700">{row['員工編號'] || row['工號'] || '-'}</span>
                            </div>
                          </td>
                          <td className="px-6 py-4">
                            <div className="flex items-center gap-2">
                              <span className="text-sm font-medium text-slate-700">{displayDate}</span>
                              {isFriday && <span className="text-[10px] font-bold text-indigo-600 bg-indigo-100 px-1.5 py-0.5 rounded">週五</span>}
                            </div>
                          </td>
                          <td className={`px-6 py-4 text-sm font-medium text-center ${mods.has('實際上班時間') ? 'bg-yellow-100 text-rose-600 font-bold' : 'text-slate-600'}`}>
                            {row['實際上班時間']}
                          </td>
                          <td className={`px-6 py-4 text-sm font-medium text-center ${mods.has('實際下班時間') ? 'bg-yellow-100 text-rose-600 font-bold' : 'text-slate-600'}`}>
                            {row['實際下班時間']}
                          </td>
                          <td className={`px-6 py-4 text-sm text-center font-bold ${mods.has('遲到(分鐘)') ? 'bg-yellow-100 text-rose-600' : 'text-slate-800'}`}>
                            {Number(row['遲到(分鐘)']) > 0 ? row['遲到(分鐘)'] : '-'}
                          </td>
                          <td className={`px-6 py-4 text-sm text-center font-bold ${mods.has('早退(分鐘)') ? 'bg-yellow-100 text-rose-600' : 'text-slate-800'}`}>
                            {Number(row['早退(分鐘)']) > 0 ? row['早退(分鐘)'] : '-'}
                          </td>
                          <td className="px-6 py-4 text-sm text-center">
                            {row['異動'] === 'V' ? (
                              <span className="px-2 py-0.5 bg-rose-100 text-rose-600 rounded font-black text-xs">V</span>
                            ) : '-'}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
