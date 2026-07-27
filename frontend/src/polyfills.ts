// Browser polyfills for exceljs (Buffer is used internally by ExcelJS)
import { Buffer } from 'buffer';
(window as any).Buffer = Buffer;
