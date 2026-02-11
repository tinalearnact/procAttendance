
export interface AttendanceRecord {
  '出勤日期': string | number | Date;
  '實際上班時間': string | number | null;
  '實際下班時間': string | number | null;
  '假勤起始時間'?: string | number | null;
  '假勤結束時間'?: string | number | null;
  '應出勤時數(時:分)'?: string | number | null;
  '遲到(分鐘)': number | string;
  '早退(分鐘)': number | string;
  '異動'?: string; // 用於標記邏輯異動 V
  '星期五'?: string; // 新增欄位：用於標記星期五 V
  _modifiedFields?: Set<string>; // 用於追蹤哪些欄位被邏輯異動過
  [key: string]: any;
}

export interface ExcelReadResult {
  data: AttendanceRecord[];
  headers: string[];
}
