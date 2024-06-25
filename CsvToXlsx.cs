using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//===================================================
// 更新履歴
//===================================================
// 2024/06/20 : 012020048D : 新規追加
// 2024/06/21 : 012020048D : Task3 : csv_to_xlsx_w_split() 追加
//===================================================


internal static class CsvToXlsx
{
    //===================================================
    // CSV→XLSX変換機能（ファイル分割あり）
    // file_path : 対象ファイルのパス
    // folder_path : 出力先フォルダのパス
    //===================================================
    static public void csv_to_xlsx_w_split(string file_path, string folder_path)
    {
        // 出力用ファイル（CSV）
        string output_csv = "分割-";
        string output_csv_name;

        // 出力用ファイル（Excel）
        string output_xlsx_name;
        string output_sheet_name = "出力シート";

        string Ex_line; // 抽出行（加工済）
        string Ex_line_mem = "dummy"; // 抽出行（前回記憶値）

        string line_1st; // 抽出行の先頭文字
        string line_1st_U = "ヘッダ";
        string line_1st_mem = "dummy"; // 持ち越し用記憶値

        string blank_line = "\0";
        string blank_data = "";
        string split_sign = ",";
        string end_line = "レポートメッセージ";

        // テスト
        int char_max = 1000;         // ファイル作成上限
        long len_max = 1000000;     // ライン上限
        long line_limit = 950000;   // ファイル切り替え閾値

        bool header_flg = false;
        int header_judge = 32;

        int row_cnt;

        // 行数を取得(ヘッダー分含まず)
        string[] lines = File.ReadAllLines(file_path, Encoding.GetEncoding("Shift_JIS"));
        long cntRow = lines.Length - 1;

        // 分割数の決定
        long file_max_L = cntRow / len_max;   // 商(切捨)
                                              //int file_max = Convert.ToInt32(file_max_L) + 1;
        int file_max = char_max;

        // ファイル読込
        using (var sr = new System.IO.StreamReader(file_path, Encoding.GetEncoding("Shift_JIS")))
        {

            // CSV作成(新規作成)
            for (int file_num = 1; file_num <= file_max; file_num++)
            {
                output_csv_name = output_csv + line_1st_U + "-" + file_num + ".csv";
                output_xlsx_name = output_csv + line_1st_U + "-" + file_num + ".xlsx";

                //出力先パスの作成
                var move_path = folder_path + @"\" + output_xlsx_name;

                using (var output_xlsx = new ExcelPackage())
                //using (var sw = new System.IO.StreamWriter(output_csv_name, false, System.Text.Encoding.GetEncoding("shift_jis")))
                {

                    // var outputSheet = output_xlsx.Workbook.Worksheets[0];
                    var outputSheet = output_xlsx.Workbook.Worksheets.Add(output_sheet_name);
                    row_cnt = 1;

                    // 前回記憶値の書き込み
                    if (Ex_line_mem != "dummy")
                    {
                        // 書き込む
                        //sw.WriteLine(Ex_line_mem);

                        // CSVデータの分割格納（lineの文字を,で区切る）
                        string[] Split_data_mem = Ex_line_mem.Split(',');

                        Ex_line_mem = "dummy";

                        // Excel書き込み
                        for (int data_num = 0; data_num < Split_data_mem.Length; data_num++)
                        {
                            outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data_mem[data_num];

                        }
                        row_cnt++;

                    }


                    // 回数分ループ
                    for (int line_cnt = 0; line_cnt < len_max; line_cnt++)
                    {
                        // 1行読み出す
                        Ex_line = sr.ReadLine();

                        if (Ex_line != null)
                        {

                            // CSVデータの分割格納（lineの文字を,で区切る）
                            string[] Split_data = Ex_line.Split(',');

                            // モニタ用
                            var split_moni_0 = Split_data[0];

                            // ヘッダ（設定オプション等）の書き込み
                            if ((line_cnt < header_judge) && (header_flg == false))
                            {
                                // 書き込む
                                //sw.WriteLine(Ex_line);

                                // Excel書き込み
                                for (int data_num = 0; data_num < Split_data.Length; data_num++)
                                {
                                    outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data[data_num];


                                }

                                row_cnt++;

                            }
                            else if (Ex_line.Contains(end_line))
                            {

                                file_max = file_num;
                                break;
                            }
                            // 本体部分の書き込み
                            else
                            {
                                // ヘッダフラグのOFF処理
                                if (header_flg == false)
                                {

                                    header_flg = true;

                                }
                                // 1列目が空白、または、空行文字のみの場合
                                if ((Split_data[0] == blank_line)
                                    || (Split_data[0] == blank_data))
                                {

                                    // 書き込む
                                    //sw.WriteLine(Ex_line);

                                    // Excel書き込み
                                    for (int data_num = 0; data_num < Split_data.Length; data_num++)
                                    {
                                        outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data[data_num];

                                    }

                                    row_cnt++;

                                }
                                //  1列目に情報がある場合
                                else if (Split_data[0] != blank_line)
                                {
                                    // 先頭文字の取得（大文字）
                                    line_1st = Convert.ToString(Ex_line.FirstOrDefault());
                                    line_1st_U = line_1st.ToUpper();


                                    // 先頭文字が前回値と一致、かつ、1ファイルの出力上限未満の場合
                                    if ((line_1st != split_sign)
                                        && (line_cnt < line_limit)
                                        && ((line_1st_U == line_1st_mem)
                                            || (line_1st_U == "dummy")))
                                    {
                                        // 書き込む
                                        //sw.WriteLine(Ex_line);

                                        // Excel書き込み
                                        for (int data_num = 0; data_num < Split_data.Length; data_num++)
                                        {
                                            outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data[data_num];

                                        }
                                        row_cnt++;
                                    }
                                    else if ((line_cnt >= line_limit)
                                            || ((line_1st_U != line_1st_mem)
                                                || (line_1st_U != blank_line)
                                                || (line_1st_U != split_sign)))
                                    {
                                        Ex_line_mem = Ex_line;
                                        line_1st_mem = line_1st_U;

                                        var fileInfo = new FileInfo(output_xlsx_name);
                                        output_xlsx.SaveAs(fileInfo);

                                        //出力ファイルの移動
                                        if (System.IO.Directory.Exists(folder_path) == true)
                                        {

                                            System.IO.File.Move(output_xlsx_name, move_path);

                                        }

                                        break;
                                    }
                                    else
                                    {
                                        // 書き込む
                                        //sw.WriteLine(Ex_line);

                                        // Excel書き込み
                                        for (int data_num = 0; data_num < Split_data.Length; data_num++)
                                        {
                                            outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data[data_num];

                                        }
                                        row_cnt++;
                                    }

                                }
                                else
                                {

                                    // 書き込む
                                    //sw.WriteLine(Ex_line);

                                    // Excel書き込み
                                    for (int data_num = 0; data_num < Split_data.Length; data_num++)
                                    {
                                        outputSheet.Cells[row_cnt, (data_num + 1)].Value = Split_data[data_num];

                                    }
                                    row_cnt++;

                                }
                            }
                        }
                        else
                        {
                            var fileInfo = new FileInfo(output_xlsx_name);
                            output_xlsx.SaveAs(fileInfo);

                            //出力ファイルの移動
                            if (System.IO.Directory.Exists(folder_path) == true)
                            {

                                System.IO.File.Move(output_xlsx_name, move_path);

                            }

                            break;
                        }
                    }

                    output_xlsx?.Dispose();
                }


            }

            sr?.Close();
        }

        return;
    }

}

