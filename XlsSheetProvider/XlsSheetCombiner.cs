using NPOI.HSSF.UserModel;
using System;
using System.IO;

namespace XlsSheetProvider
{
    /*
    * 使用 NPOI 套件將多個
    *  指定索引 EXCEL (*.xls) 的 Sheet (工作表) 合併成新的 EXCEL(*.xls)
    *  
    * 此專案為 NET.Standard
    *  僅適用於 full .NET Framework 版本
    */

    /// <summary>提供進行 EXCEL(*.xls) 工作表合併方法</summary>
    public sealed class XlsSheetCombiner
    {
        private HSSFWorkbook _workbook;
        private Byte[] _buffer;


        public XlsSheetCombiner()
        {
            this._workbook = new HSSFWorkbook();
        }

        /// <summary>加入指定 EXCEL (*.xls) 檔案指定的索引 Sheet</summary>
        /// <param name="buffer">檔案位元陣列</param>
        /// <param name="sheetIndex">複製的工作表索引</param>
        /// <param name="sheetName">工作表名稱</param>
        public void Add(byte[] buffer, int sheetIndex, string sheetName)
        {
            using (MemoryStream _stream = new MemoryStream(buffer))
            {
                HSSFWorkbook _workbook = new HSSFWorkbook(_stream);
                HSSFSheet _sheet = _workbook.GetSheetAt(sheetIndex) as HSSFSheet;

                _sheet.CopyTo(this._workbook, sheetName, true, true);
            }
        }

        /// <summary>儲存工作表設定</summary>
        public void Save()
        {
            using (MemoryStream _stream = new MemoryStream())
            {
                this._workbook.Write(_stream);
                this._buffer = _stream.ToArray();
            }
        }

        /// <summary>取得合併後的 EXCEL 檔案位元陣列（唯讀）</summary>
        public byte[] XlsContent
        {
            get
            {
                return this._buffer;
            }
        }


    }
}
