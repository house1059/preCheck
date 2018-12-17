using Microsoft.VisualBasic;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using SharpSvn;

namespace DataConvertCheckTool {

enum ErrorCode :int
{
    eTableName,     //テーブル名が同じ
}



    class Program {




        //summary
        //引数：チェックしたいパスが記載されているTextファイルのパス
        static void Main(string[] args)
        {

            XlsPath xlp = new XlsPath();
            string filePath = @"C:\NSS\W119\sub\NSS_W115\M32C_8B_W115\Tool\batch\W119_データ変換ツール_Ver210a_W115枠用_jenkins.xlsm";        //quated
            //string filePath = @"C:/Users/house/OneDrive/ドキュメント/GitHub/pre/DataConvertCheckTool/data/test.xlsx";
            string[] strPath;





            xlp.DataConvertToText(filePath);

            try
            {
                //Pathファイルを作成
                StreamReader stream = new StreamReader(filePath + ".txt", Encoding.GetEncoding("shift_jis"));
                strPath = stream.ReadToEnd().Split(new[] { "\r\n" }, StringSplitOptions.None);
                stream.Close();     //ファイルioは素早く終わらせる


                //Pathファイルを読み込み簡易チェックを行う
                for(int i = 39; i < strPath.Length; i++) {      //ここの39はw119が39列目まで振分けが無いので39で固定しているだけです。
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
    Dictionary<int, string> tbNmae = new Dictionary<int, string> { };       //振分けテーブルの名称チェック
    List<ErrData> errorList = new List<ErrData>();                          //エラーリスト

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

            //svnクライアント
            SvnClient client = new SvnClient();
            
            //クライアントのファイル位置を設定  
            SvnPathTarget local = new SvnPathTarget(filePath);

            //ファイルのsvnプロパティを抜く
            SvnInfoEventArgs clientInfo;
            
            client.GetInfo(local, out clientInfo);

            foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
            {

                string s = Strings.StrConv(ws.Name, VbStrConv.Wide | VbStrConv.Uppercase);

                if (s.StartsWith("ＴＢ＿"))
                {


                    //◎の検索（なければリターン）、移動
                    //Findメソッドがないので、cellデータを取得してLinqによりアドレスを算出する
                    var query = from cell in ws.Cells where cell.Text == "◎" select cell;


                    //◎のリストが完成したのでそれぞれでデータチェック
                    if (0 < query.Count()) {
                        foreach( ExcelRangeBase range in query.ToList())
                        {
                            //振分けテーブル名の被りチェック( Dictionaryチェック）
                            if (tbNmae.ContainsValue( range.Offset(0,1).Text))
                            {
                                ErrData d = new ErrData();
                                d.Auther = clientInfo.LastChangeAuthor;
                                d.ErrCode = ErrorCode.eTableName;
                                d.ErrName = range.Offset(0, 1).Text;
                                errorList.Add(d);

                            } else
                            {
                                tbNmae.Add(tbNmae.Count + 1, range.Offset(0, 1).Text);
                            }

                            //振分け合計値の確認+データ有りの斜線チェック（W119では0データに斜線はＯＫ）
                            int dataMax = int.Parse(range.Offset(3, 1).Text);


                            //データ項目列（標準）まで移動　右移動で数値になるまで
                            int dataEndColumn = 0;
                            int dataEndRow = 0;

                            ExcelRangeBase exStart;
                            ExcelRangeBase exEnd;

                            //データのスタート位置
                            for (int i = 1; ; i++){
                                if(int.TryParse(range.Offset(4, 1).Offset(0, i).Text,out int result) == false)
                                {
                                    exStart = range.Offset(5, 1).Offset(0, i);
                                    break;
                                }
                            }

                            //データの横終了位置
                            for (int i = 1; ; i++)
                            {
                                if (exStart.Offset(0, i).Text == "")
                                {
                                    dataEndColumn = exStart.Offset(0, i-1).Columns;
                                    break;
                                }
                            }

                            //データの縦終了位置
                            for (int i = 1; ; i++)
                            {
                                if (exStart.Offset(i, 0).Text == "")
                                {
                                    dataEndRow = exStart.Offset( i - 1,0).Rows;
                                    break;
                                }
                            }
                            exEnd = exStart.Offset(dataEndRow, dataEndColumn);

                            ExcelRangeBase dataCheck = exStart;
                            for(int i = 1; i <= dataEndRow; i++)
                            {
                                //横に足し算しながら進めていく
                                int sum = 0;
                                if(int.TryParse( dataCheck.Offset(0,i).Value.ToString(), out int result))
                                {
                                    sum += int.Parse(dataCheck.Offset(0, i).Value.ToString());
                                }

                            }


                        }
                    }



                    



                    //振分け最大値の取得



                    //データ項目列へ移動


                    //データを縦になめていく
                    //合計値、斜線チェック

                }

            }
        }
    }


    //sammary
    //データデータを集めてこれをリスト化して出力する
    class ErrData {
        public ErrorCode    ErrCode{ get; set; }
        public string ErrName{ get; set; }
        public string Auther{ get; set; }
    }



}
