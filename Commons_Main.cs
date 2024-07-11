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
// 2024/06/15 : 012020048D : Task1 : File_Select() 追加
// 2024/06/16 : 012020048D : Task1 : Folder_Select()追加
// 2024/06/22 : 012020048D : Task1 : WorkFolder_Create()追加
// 2024/06/20 : 012020048D : Task2 : CallMacro()追加
//===================================================

public class Commons
{
	//===================================================
	// ファイル選択
	//===================================================
	// pop_msg : 表示させるメッセージ
	// top_msg : メッセージボックスのタイトル
	// file_Name：デフォルトのファイル名
	// file_filter：[ファイルの種類]に表示される選択肢、ファイル形式
	// title_msg：ウィンドウのタイトル
	// output_file：開いたファイル
	//===================================================
	static public string File_Select(string pop_msg, string top_msg, string file_Name, string file_filter, string title_msg)
	{
		string err_txt;

		//OpenFileDialogクラスのインスタンスを作成
		System.Windows.Forms.OpenFileDialog open_ofd = new System.Windows.Forms.OpenFileDialog();

		//はじめのファイル名を指定する
		//はじめに「ファイル名」で表示される文字列を指定する
		open_ofd.FileName = file_Name;
		//はじめに表示されるフォルダを指定する
		//指定しない（空の文字列）の時は、現在のディレクトリが表示される
		//ofd.InitialDirectory = @"C:\";
		//[ファイルの種類]に表示される選択肢を指定する
		//指定しないとすべてのファイルが表示される
		open_ofd.Filter = file_filter;
		//[ファイルの種類]ではじめに選択されるものを指定する
		//2番目の「すべてのファイル」が選択されているようにする
		open_ofd.FilterIndex = 1;
		//タイトルを設定する
		open_ofd.Title = title_msg;
		//ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
		open_ofd.RestoreDirectory = true;
		//存在しないファイルの名前が指定されたとき警告を表示する
		//デフォルトでTrueなので指定する必要はない
		open_ofd.CheckFileExists = true;
		//存在しないパスが指定されたとき警告を表示する
		//デフォルトでTrueなので指定する必要はない
		open_ofd.CheckPathExists = true;

		//ダイアログを表示する
		if (open_ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
		{

			//OKボタンがクリックされたとき、選択されたファイル名を表示する
			output_file = open_ofd.FileName;
			return output_file;

		}
		else
		{
			//「いいえ」が選択された時
			err_txt = "操作がキャンセルされました。";
			MessageBox.Show(err_txt);
			return null;
		}

	}

	//===================================================
	// フォルダ選択
	// pop_msg : 表示させるメッセージ
	//===================================================
	static public string Folder_Select(string pop_msg)
	{
		string err_txt;
		string output_file;

		// 出力テンプレートを開く
		FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

		// ダイアログの説明文
		folderBrowserDialog.Description = pop_msg;


		// 「新しいフォルダーの作成する」ボタンを表示する（デフォルトはtrue）
		folderBrowserDialog.ShowNewFolderButton = false;

		//ダイアログを表示する
		if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
		{

			//OKボタンがクリックされたとき、選択されたファイル名を表示する
			output_file = folderBrowserDialog.SelectedPath;
			return output_file;

		}
		else
		{
			//「いいえ」が選択された時
			err_txt = "操作がキャンセルされました。";
			MessageBox.Show(err_txt);
			return null;
		}

	}

	//===================================================
	// 作業フォルダ作成
	// str_name : 作成するフォルダ名
	//===================================================
	public static string WorkFolder_Create(string str_folder_name)
	{
		string direct_path;
		string return_path;

		//フォルダ作成用の日付を取得する
		var output_folder_date = DateTime.Now;
		string folder_date = output_folder_date.ToString("yyyyMMdd");
		string folder_time = output_folder_date.ToString("HHmmss");
		string folder_name = str_dir_name + folder_date + folder_time;

		//出力用フォルダの作成
		System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(System.IO.Directory.GetCurrentDirectory());
		//var CurrentDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
		di.CreateSubdirectory(folder_name);

		direct_path = Convert.ToString(di);

		return_path = direct_path + @"\" + folder_name;

		return return_path;
	}

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

