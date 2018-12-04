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

            int fileCount = 0;
            
            XlsPath xlp = new XlsPath();

            //xlp.DataConvertToText("C:/NSS/W119/sub/NSS_W115/M32C_8B_W115/Tool/batch/W119_データ変換ツール_Ver210a_W115枠用_jenkins.xlsm");
            xlp.DataConvertToText("C:/Users/house/OneDrive/ドキュメント/GitHub/pre/DataConvertCheckTool/data/test.xlsx");
            
            //テキストを読み込む（本当はデータ変換ツールが通る保障まで行きたいが、データ変換ツールを通せばいいのでW119用としてとりあえず振分け表のチェックを行えればよい）
            List<String> path;
            string[] strPath;

            try
            {
                StreamReader stream = new StreamReader(args[0], Encoding.GetEncoding("shift_jis"));
                strPath =  stream.ReadToEnd().Split('\n');
                stream.Close();     //ファイルioは素早く終わらせる

            }catch (Exception e)
            {

                Console.WriteLine(e.Message);
                return;
            }
            
            
            //ループ処理
            foreach(string s in strPath)
            {
                //空白の場合無視
                if (s == "") continue;

                //パスが存在しない場合その旨をtextファイルに吐き出す
                if( File.Exists(s) == false)
                {
                    //エラーの書き出し
                }
                else
                {





                }




            }





        }
    }


    

    class XlsPath {


        StreamWriter stream;


        //summary
        //データ変換ツールのデータからフルパスのtxtを出力する
     public void DataConvertToText( string filePath  ){

            try
            {
                //Excelを開く   cellsのvalue(size,size)で最後の行が分かる？
                ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));
                ExcelWorksheet sheet = excel.Workbook.Worksheets["変換設定"];

                //出力ファイル
                int lastRow = sheet.Dimension.End.Row;
                int lastColumn = sheet.Dimension.End.Column;

                const int ROWOFFSET = 12;


                stream = new StreamWriter(filePath + ".txt", false, Encoding.GetEncoding("shift_jis"));
                stream.WriteLine("[HEAD]");
                stream.WriteLine(DateTime.Now);
                stream.WriteLine();

                for (int i = 0; i < (lastRow- ROWOFFSET+1) ; i++) {
                    ExcelRangeBase rangeBase = sheet.Cells[ROWOFFSET, 3].Offset(i, 0);
                    if (null!= rangeBase.Value && rangeBase.Value.ToString() == "変換ファイル名（フルパス）"){
                        if (null != rangeBase.Offset(0, 1).Value)
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
    }





    public void CheckSheet(string filePath)
        {
            //Excelを開く
            ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));

            foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
            {
                //tb_　シートのみを対象とする
                if (Strings.StrConv(ws.Name, VbStrConv.Wide & VbStrConv.Uppercase).StartsWith("ＴＢ＿"))
                {

                    //◎の検索（なければリターン）、移動



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
