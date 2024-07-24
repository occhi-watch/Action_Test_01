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
// 2024/07/24 : 012020048D : 新規追加
// 2024/07/24 : 012020048D : Convert_Dot_to_Text()追加
//                         : ContrlFlow_Convertt()追加
//===================================================

public class DotToText
{
    //===================================================
    // 共通変数
    //===================================================
    public static string str_rename_end_Dot = @".Dot";
    public static string str_rename_end_dot = @".dot";
    public static string str_rename_end_txt = @".txt";
    //===================================================

    //===================================================
    // Dot → txt 変換
    //===================================================
    public static void ContrlFlow_Convert(string str_dot_folder)
    {

        // カレントディレクトリを変更
        Environment.CurrentDirectory = Directory.GetCurrentDirectory();

        //===================================================
        // Dot → txt 変換(大文字)
        //===================================================

        // フォルダ内のDotファイルを取得(大文字)
        string[] Dot_file_List = Directory.GetFiles(str_dot_folder, str_rename_end_Dot);

        foreach (var file_path in Dot_file_List)
        {
            if (file_path.Contains(str_rename_end_Dot))
            {
                Convert_Dot_to_Text(file_path, str_rename_end_Dot);
            }
            else
            {
                // 何もしない
            }
        }

        //===================================================
        // Dot → txt 変換(小文字)
        //===================================================

        // フォルダ内のdotファイルを取得(小文字)
        string[] dot_file_List = Directory.GetFiles(str_dot_folder, str_rename_end_dot);

        foreach (var file_path in dot_file_List)
        {

            if (file_path.Contains(str_rename_end_dot))
            {
                Convert_Dot_to_Text(file_path, str_rename_end_dot);
            }
            else
            {
                // 何もしない
            }
        }


    }

    //===================================================
    // Dot→txt変換
    //===================================================
    // str_direct_path  : 対象フォルダの絶対パス
    // convert_end      : 対象となる拡張子
    //===================================================
    public static void Convert_Dot_to_Text(string str_direct_path, string convert_end)
    {
        // 変更後の文字列を作成（Dot→txt）
        var rename_filepath = str_direct_path.Replace(convert_end, str_rename_end_txt);

        // パスからファイル名を取得
        var filename = Path.GetFileName(rename_filepath);

        //出力先パスの作成
        var move_path = str_dot_folder + @"\" + filename;

        //出力ファイルのコピー
        System.IO.File.Copy(str_direct_path, move_path);

    }

}

