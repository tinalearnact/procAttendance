
import XLSX from 'xlsx';
import { AttendanceRecord, ExcelReadResult } from '../types';

/**
 * 將時間格式轉換為當日分鐘數
 */
const timeToMinutes = (val: any): number | null => {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  
  if (typeof val === 'string') {
    const cleanStr = val.trim();
    if (cleanStr === '(未打卡)') return null;
    const parts = cleanStr.split(':');
    if (parts.length >= 2) {
      const h = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      return isNaN(h) || isNaN(m) ? null : h * 60 + m;
    }
  }

  if (typeof val === 'number') {
    return Math.round(val * 24 * 60) % 1440;
  }
  return null;
};

const isEmpty = (val: any): boolean => {
  return val === null || val === undefined || String(val).trim() === '';
};

const parseDate = (val: any): Date => {
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(Math.round((val - 25569) * 864e5));
  return new Date(String(val));
};

export const processAttendanceData = (data: AttendanceRecord[]): AttendanceRecord[] => {
  return data.map((row) => {
    const dateObj = parseDate(row['出勤日期']);
    const isFri = !isNaN(dateObj.getTime()) && dateObj.getDay() === 5;
    
    const newRow = { ...row };
    // 1. 星期五標記邏輯
    newRow['星期五'] = isFri ? 'V' : '';
    
    // 僅對星期五進行異動邏輯處理
    if (!isFri) {
      newRow['異動'] = '';
      newRow._modifiedFields = new Set();
      return newRow;
    }

    const modifiedFields = new Set<string>();

    const checkInVal = newRow['實際上班時間'];
    const checkOutVal = newRow['實際下班時間'];
    const leaveStartVal = newRow['假勤起始時間'];
    const leaveEndVal = newRow['假勤結束時間'];
    const scheduledHoursVal = newRow['應出勤時數(時:分)'];
    
    const checkInMin = timeToMinutes(checkInVal);
    const checkOutMin = timeToMinutes(checkOutVal);
    const leaveStartMin = timeToMinutes(leaveStartVal);
    const scheduledMin = timeToMinutes(scheduledHoursVal);

    // 判斷遲到基準：07:00 -> 10:00 (600min), else 09:00 (540min)
    const lateThreshold = (scheduledMin === 420) ? 600 : 540;

    // 上班處理
    if (isEmpty(checkInVal) && isEmpty(leaveStartVal)) {
      newRow['實際上班時間'] = '(未打卡)';
      modifiedFields.add('實際上班時間');
    } else if (checkInMin !== null && checkInMin > lateThreshold) {
      // 邏輯優化：計算出遲到分鐘大於 0 時，當 假勤起始時間 早於 實際上班時間，表示有請假，不用異動遲到欄位
      const hasValidLeave = leaveStartMin !== null && leaveStartMin < checkInMin;
      if (!hasValidLeave) {
        newRow['遲到(分鐘)'] = Math.floor(checkInMin - lateThreshold);
        modifiedFields.add('遲到(分鐘)');
      }
    }

    // 下班處理
    if (isEmpty(checkOutVal) && isEmpty(leaveEndVal)) {
      newRow['實際下班時間'] = '(未打卡)';
      modifiedFields.add('實際下班時間');
    } else if (checkOutMin !== null && checkOutMin < 1080) { // 18:00
      // 邏輯優化：計算出早退分鐘大於 0 時，當 假勤起始時間 早於 實際下班時間，表示有請假，不用異動早退欄位
      const hasValidLeave = leaveStartMin !== null && leaveStartMin < checkOutMin;
      if (!hasValidLeave) {
        newRow['早退(分鐘)'] = Math.floor(1080 - checkOutMin);
        modifiedFields.add('早退(分鐘)');
      }
    }

    // 如果有異動，填入 V
    if (modifiedFields.size > 0) {
      newRow['異動'] = 'V';
    } else {
      newRow['異動'] = '';
    }

    newRow._modifiedFields = modifiedFields;
    return newRow;
  });
};

/**
 * 讀取 Excel 並保留標題順序
 */
export const readExcel = (file: File): Promise<ExcelReadResult> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // 取得原始標題順序
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        const headers: string[] = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
          headers.push(cell ? String(cell.v) : `Unknown_${C}`);
        }

        const jsonData = XLSX.utils.sheet_to_json<AttendanceRecord>(worksheet);
        
        const skipKeywords = ['小計', '合計', '總計', '次數', '遲到早退'];
        const validData = jsonData.filter(row => {
          const d = row['出勤日期'];
          if (!d) return false;
          if (skipKeywords.some(key => String(d).includes(key))) return false;
          if (isNaN(parseDate(d).getTime())) return false;
          return true;
        });
        
        resolve({ data: validData, headers });
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

/**
 * 匯出 Excel 並確保黃底紅字樣式、欄位順序、以及新增「異動」與「星期五」欄位
 */
export const exportToExcel = (data: AttendanceRecord[], originalHeaders: string[], fileName: string) => {
  if (data.length === 0) return;

  // 1. 確保標題包含「異動」與「星期五」且位於最後
  let headers = [...originalHeaders];
  if (!headers.includes('異動')) headers.push('異動');
  if (!headers.includes('星期五')) headers.push('星期五');

  // 2. 準備匯出資料
  const exportRows = data.map(row => {
    const formattedRow: any = {};
    headers.forEach(h => {
      const val = row[h];
      if (h === '出勤日期' && val instanceof Date) {
        formattedRow[h] = `${val.getFullYear()}/${(val.getMonth() + 1).toString().padStart(2, '0')}/${val.getDate().toString().padStart(2, '0')}`;
      } else {
        formattedRow[h] = (val === undefined || val === null) ? '' : val;
      }
    });
    return formattedRow;
  });

  // 3. 建立工作表
  const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: headers });

  // 4. 套用黃底紅字樣式到「異動儲存格」
  data.forEach((row, rowIndex) => {
    if (row._modifiedFields && row._modifiedFields.size > 0) {
      headers.forEach((header, colIndex) => {
        if (row._modifiedFields?.has(header)) {
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: colIndex });
          if (worksheet[cellAddress]) {
            worksheet[cellAddress].s = {
              fill: {
                patternType: 'solid',
                fgColor: { rgb: "FFFF00" } 
              },
              font: {
                name: 'Arial',
                sz: 11,
                bold: true,
                color: { rgb: "FF0000" } 
              },
              alignment: {
                horizontal: 'center',
                vertical: 'center'
              },
              border: {
                top: { style: 'thin', color: { rgb: "000000" } },
                bottom: { style: 'thin', color: { rgb: "000000" } },
                left: { style: 'thin', color: { rgb: "000000" } },
                right: { style: 'thin', color: { rgb: "000000" } }
              }
            };
          }
        }
      });
    }
  });

  // 設定自動欄寬
  worksheet['!cols'] = headers.map((h) => ({ wch: (h === '異動' || h === '星期五') ? 8 : 15 }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '考勤處理結果');
  
  // 5. 寫入檔案 (.xlsx)
  const baseName = fileName.replace(/\.[^/.]+$/, "");
  XLSX.writeFile(workbook, `邏輯處理_${baseName}.xlsx`);
};
