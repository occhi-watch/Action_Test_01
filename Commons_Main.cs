using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

//===================================================
// 更新履歴
//===================================================
// 2024/06/11 : 012020048D : 新規追加
// 2024/06/20 : 012020048D : Task2 : CallMacro()追加
//===================================================

public class Commons
{
    //===================================================
    // ExcelVBA マクロ呼び出し
    // str_path : 呼び出すxlsmファイルのパス
    // str_macro : 呼び出すExcelVBAマクロ名
    // row_cnt : 対象の行数
    // col_cnt : 対象の列数
    // str_direct_path : 出力先の絶対パス
    //===================================================
    public static void CallMacro(string str_xlsm_path, string str_macro, int row_cnt, int col_cnt, string str_direct_path)
    {
        // Excel.Application の新しいインスタンスを生成する
        var xlApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbooks xlBooks;

        // xlApplication から WorkBooks を取得する
        // 既存の Excel ブックを開く
        var CurrentDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

        xlBooks = xlApp.Workbooks;

        if (System.IO.File.Exists(str_xlsm_path) == true)
        {
            // ブックを開く
            xlBooks.Open(str_xlsm_path);

            // Excel を表示する
            xlApp.Visible = false;

            // マクロを実行する
            xlApp.Run(str_macro, row_cnt, col_cnt, str_direct_path);

            // Excel を終了する
            xlBooks.Close();
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }
        else
        {
            // 存在しない
            var Text = "ExcelVBAが実行出来ませんでした。";
            MessageBox.Show(Text);
            return;
        }

    }

}

