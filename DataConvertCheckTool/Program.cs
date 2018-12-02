﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.VisualBasic;
using System.IO;
using ClosedXML.Excel;


namespace DataConvertCheckTool {
    class Program {


        //summary
        //引数：チェックしたいパスが記載されているTextファイルのパス
        static void Main(string[] args)
        {

            int fileCount = 0;
            
            XlsPath xlp = null;

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
                //Excelを開く
                ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));
                ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets["変換設定"];

                //ClosedXml
                XLWorkbook xLWorkbook = new XLWorkbook(filePath);
                IXLWorksheet workSheet = xLWorkbook.Worksheet("変換設定");


                //出力ファイル
                stream = new StreamWriter(filePath, false, Encoding.GetEncoding("shift_jis"));
                stream.WriteLine("[HEAD]");
                stream.WriteLine(DateTime.Now);
                stream.WriteLine();


                //データ変換ツールに記載のパスを取得
                //最終データ行を取得する場合はepplusを使うより、ClosedXmlを使う早い
                var lastRow = workSheet.RangeUsed.ragen
                    workSheet.rangeUsed.RangeAddress.LastAddress.RowNumber


            }
            catch (Exception e)
            {








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
