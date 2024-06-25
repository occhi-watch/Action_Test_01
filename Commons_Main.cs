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
    //===================================================
    public static void CallMacro(String path, String macro, int ap_cnt, int var_cnt, string direct_path)
    {
        // Excel.Application の新しいインスタンスを生成する
        var xlApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbooks xlBooks;

        //MessageBox.Show(direct_path);
        // xlApplication から WorkBooks を取得する
        // 既存の Excel ブックを開く
        var CurrentDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        //path = CurrentDirectory + @"\" + path;
        xlBooks = xlApp.Workbooks;

        if (System.IO.File.Exists(path) == true)
        {

            xlBooks.Open(path);
            // Excel を表示する
            xlApp.Visible = false;

            // マクロを実行する
            // 標準モジュール内のTestメソッドに "Hello World" を引数で渡し実行
            //xlApp.Run("work.xlsm!Test", "Hello World");
            // Sheet1内のSheetTestメソッドを実行(引数なし)
            xlApp.Run(macro, ap_cnt, var_cnt, direct_path);

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

