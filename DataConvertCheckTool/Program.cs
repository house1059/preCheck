using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.VisualBasic;
using ClosedXML.Excel;


namespace DataConvertCheckTool {
    class Program {


        //summary
        //引数：チェックしたいパスが記載されているTextファイルのパス
        static void Main(string[] args)
        {

            XlsPath xlp = new XlsPath();
            //string filePath = "C:/NSS/W119/sub/NSS_W115/M32C_8B_W115/Tool/batch/W119_データ変換ツール_Ver210a_W115枠用_jenkins.xlsx";
            string filePath = "C:/Users/house/OneDrive/ドキュメント/GitHub/pre/DataConvertCheckTool/data/test.xlsx";
            string[] strPath;


            xlp.DataConvertToText(filePath);
            
            try
            {
                //Pathファイルを作成
                StreamReader stream = new StreamReader(filePath + ".txt", Encoding.GetEncoding("shift_jis"));
                strPath =  stream.ReadToEnd().Split('\n');
                stream.Close();     //ファイルioは素早く終わらせる


                //Pathファイルを読み込み簡易チェックを行う
                for(int i = 3; i < strPath.Length; i++) { 
                    //仕様書の各ファイルを開いて簡易チェックを行う
                    xlp.CheckSheet(strPath[i]);
                }

            }catch (Exception e)
            {

                Console.WriteLine(e.Message);
                return;
            }
       
        }
    }
    

class XlsPath {

    StreamWriter stream;

    //summary
    //データ変換ツールのデータからフルパスのtxtを出力する
    public string DataConvertToText( string filePath  ){

            string txtPath = filePath + ".txt";
            try
            {
                //Excelを開く   cellsのvalue(size,size)で最後の行が分かる？
                ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));
                ExcelWorksheet sheet = excel.Workbook.Worksheets["変換設定"];

                //出力ファイル
                int lastRow = sheet.Dimension.End.Row;
                int lastColumn = sheet.Dimension.End.Column;


                stream = new StreamWriter(txtPath, false, Encoding.GetEncoding("shift_jis"));
                stream.WriteLine("[HEAD]");
                stream.WriteLine(DateTime.Now);
                stream.WriteLine();

                for (int i = 0; i <= (lastRow) ; i++) {
                    ExcelRangeBase rangeBase = sheet.Cells[1, 3].Offset(i, 0);
                    if (null!= rangeBase.Value && rangeBase.Value.ToString() == "変換ファイル名（フルパス）"){
                        if (null != rangeBase.Offset(0, 1).Value && rangeBase.Offset(0, 1).Value.ToString() != "")
                        {
                            stream.WriteLine(rangeBase.Offset(0, 1).Value.ToString());
                        }
                    }
                }
                stream.Close();
            }
            catch (Exception e)
            {
                StreamWriter error = new StreamWriter(filePath + ".log", false, Encoding.GetEncoding("shift_jis"));
                error.WriteLine("Error log");
                error.WriteLine(e.Message);
                error.WriteLine(e.Source);
                error.Close();
            }
            return txtPath;
    }

    //summary
    //指定したファイルを開き簡易チェックを行う。
    public void CheckSheet(string filePath)
    {

            //fileExists
            if (File.Exists(filePath) == false) return;

            //Excelを開く
            ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));

            foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
            {
                //tb_　シートのみを対象とする
                if (Strings.StrConv(ws.Name, VbStrConv.Wide & VbStrConv.Uppercase).StartsWith("ＴＢ＿"))
                {

                    //◎の検索（なければリターン）、移動
                    //Findメソッドがないので、cellデータを取得してLinqによりアドレスを算出する
                    var find = from cell in ws.Cells where cell.Text == "◎" select cell;


                    //振分けテーブル名の被りチェック( Dictionaryチェック）
                    //→被っていれば被りエラーをｺﾚｸｼｮﾝ                    


                    //振分け最大値の取得



                    //データ項目列へ移動


                    //データを縦になめていく
                    //合計値、斜線チェック

                }

            }
        }
    }




}
