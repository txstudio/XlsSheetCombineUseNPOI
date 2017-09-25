using System;
using System.IO;
using XlsSheetProvider;

namespace ConsoleApp
{
    class Program
    {
        const string _stock_etf_path = "../../file/stock_etf.xls";
        const string _stock_list_path = "../../file/stock_list.xls";
        const string _stock_list_2_path = "../../file/stock_list_2.xls";

        const string _outPath = "../../file/out.xls";

        static void Main(string[] args)
        {
            byte[] _etfBuffer;
            byte[] _listBuffer;
            byte[] _list2Buffer;
            byte[] _outBuffer;

            XlsSheetCombiner _combiner;

            _combiner = new XlsSheetCombiner();


            //取得要合併的 EXCEL 檔案
            _etfBuffer = File.ReadAllBytes(_stock_etf_path);
            _listBuffer = File.ReadAllBytes(_stock_list_path);
            _list2Buffer = File.ReadAllBytes(_stock_list_2_path);


            //進行工作表合併
            _combiner.Add(_etfBuffer, 0, "ETF");
            _combiner.Add(_listBuffer, 0, "上市櫃公司一覽表");
            _combiner.Add(_list2Buffer, 0, "興貴公司一覽表");
            _combiner.Save();


            //取得合併後的檔案位元陣列
            _outBuffer = _combiner.XlsContent;

            
            File.WriteAllBytes(_outPath, _outBuffer);

            
            Console.WriteLine("press any key to exit");
            Console.ReadKey();
        }
    }
}
