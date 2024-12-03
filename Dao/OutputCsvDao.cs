using ExcelConvertToOkumarukunnCsv.Dto;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelConvertToOkumarukunnCsv.Dao
{
    public class OutputCsvDao
    {
        const string csvHeader00 = "\"ディーラーコード\"";
        const string csvHeader01 = "\"運送会社コード\"";
        const string csvHeader02 = "\"運送会社名\"";
        const string csvHeader03 = "\"送り状No\"";
        const string csvHeader04 = "\"出荷日\"";
        const string csvHeader05 = "\"出荷元コード\"";
        const string csvHeader06 = "\"出荷元名\"";
        const string csvHeader07 = "\"発送先住所\"";
        const string csvHeader08 = "\"発送先コード\"";
        const string csvHeader09 = "\"発送先名\"";
        const string csvHeader10 = "\"品目コード\"";
        const string csvHeader11 = "\"JANコード\"";
        const string csvHeader12 = "\"品名\"";
        const string csvHeader13 = "\"品番\"";
        const string csvHeader14 = "\"ロットNo\"";
        const string csvHeader15 = "\"納品書No\"";
        const string csvHeader16 = "\"出荷数\"";
        const string csvHeader17 = "\"備考\"";
        const string csvHeader18 = "\"担当営業所コード\"";
        const string csvHeader19 = "\"担当営業所名\"";
        const string csvHeader20 = "\"受注日\"";
        const string csvHeader21 = "\"ディーラー発注番号\"";
        const string csvHeader22 = "\"病院名\"";
        const string csvHeader23 = "\"受注No\"";
        const string csvHeader24 = "\"シリアルNo\"";
        const string csvHeader25 = "\"使用期限\"";
        const string csvHeader26 = "\"症例日\"";
        const string csvHeader27 = "\"オーダー種類\"";
        const string csvHeader28 = "\"事業部名\"";
        const string csvHeader29 = "\"単位名\"";

        public void Output(string filePath, List<PropertyDto> dataList)
        {
            IList<PropertyDto> CsvData = new List<PropertyDto>();
            DataTable dt = new DataTable();
            var properties = typeof(PropertyDto).GetProperties();

            string csvPath = System.IO.Path.ChangeExtension(filePath, "csv");

            // ヘッダ定数を取得し配列化
            string[] HeadArray = { csvHeader00, csvHeader01, csvHeader02, csvHeader03, csvHeader04, csvHeader05, csvHeader06,
                                   csvHeader07, csvHeader08, csvHeader09, csvHeader10, csvHeader11, csvHeader12, csvHeader13,
                                   csvHeader14, csvHeader15, csvHeader16, csvHeader17, csvHeader18, csvHeader19, csvHeader20,
                                   csvHeader21, csvHeader22, csvHeader23, csvHeader24, csvHeader25, csvHeader26, csvHeader27,
                                   csvHeader28, csvHeader29 };

            // ヘッダを格納
            for (int i = 0; i < HeadArray.Length; i++)
                dt.Columns.Add(HeadArray[i]);

            int cnt = 0;
            // レコードを格納
            foreach (var item in dataList)
            {
                var row = dt.NewRow();
                foreach (var prop in properties)
                {
                    var itemValue = prop.GetValue(item, new object[] { });
                    row[HeadArray[cnt]] = itemValue;
                    cnt++;
                }
                dt.Rows.Add(row);
                cnt = 0;
            }

            // CSVファイルに書き込む
            using (var writer = new System.IO.StreamWriter(csvPath, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                //ヘッダを書き込む
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    //取得
                    string col = dt.Columns[i].Caption;
                    writer.Write(col);
                    //カンマを書き込む
                    if (dt.Columns.Count - 1 > i)
                    {
                        writer.Write(',');
                    }
                }
                //改行する
                writer.Write("\r\n");

                //レコードを書き込む
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        //取得
                        string col = row[i].ToString();
                        //ダブルクォーテーションで囲む
                        col = EncloseDoubleQuotes(col);
                        //書き込む
                        writer.Write(col);
                        //カンマを書き込む
                        if (dt.Columns.Count - 1 > i)
                        {
                            writer.Write(',');
                        }
                    }
                    writer.Write("\r\n");
                }

                //閉じる
                writer.Close();
            }
        }

        private string EncloseDoubleQuotes(string value)
        {
            return "\"" + value.Replace("\"", "\"\"") + "\""; // ダブルクォーテーションを適切にエスケープ
        }
    }

    public class OutputCsvDaoJ
    {
        const string csvHeader00 = "\"ディーラーコード\"";
        const string csvHeader01 = "\"データ区分\"";
        const string csvHeader02 = "\"受注No\"";
        const string csvHeader03 = "\"受注日\"";
        const string csvHeader04 = "\"送り状No\"";
        const string csvHeader05 = "\"発送先名\"";
        const string csvHeader06 = "\"ディーラー発注番号\"";
        const string csvHeader07 = "\"発送先住所\"";
        const string csvHeader08 = "\"品目コード\"";
        const string csvHeader09 = "\"JANコード\"";
        const string csvHeader10 = "\"品名\"";
        const string csvHeader11 = "\"品番\"";
        const string csvHeader12 = "\"ロットNo\"";
        const string csvHeader13 = "\"シリアルNo\"";
        const string csvHeader14 = "\"使用期限\"";
        const string csvHeader15 = "\"納品書No\"";
        const string csvHeader16 = "\"受注数\"";
        const string csvHeader17 = "\"出荷数\"";
        const string csvHeader18 = "\"単位名\"";
        const string csvHeader19 = "\"出荷日\"";
        const string csvHeader20 = "\"病院名\"";
        const string csvHeader21 = "\"担当営業所名\"";
        const string csvHeader22 = "\"出荷元名\"";
        const string csvHeader23 = "\"運送会社コード\"";
        const string csvHeader24 = "\"運送会社名\"";
        const string csvHeader25 = "\"備考\"";
        const string csvHeader26 = "\"症例日\"";
        const string csvHeader27 = "\"オーダー種類\"";
        const string csvHeader28 = "\"事業部名\"";
        const string csvHeader29 = "\"メール通知区分\"";
        const string csvHeader30 = "\"発送先コード\"";

        public void OutputJ(string filePath, List<PropertyDtoJ> dataList)
        {
            IList<PropertyDtoJ> CsvData = new List<PropertyDtoJ>();
            DataTable dt = new DataTable();
            var properties = typeof(PropertyDtoJ).GetProperties();

            string csvPath = System.IO.Path.ChangeExtension(filePath, "csv");

            // ヘッダ定数を取得し配列化
            string[] HeadArray = { csvHeader00, csvHeader01, csvHeader02, csvHeader03, csvHeader04, csvHeader05, csvHeader06,
                                   csvHeader07, csvHeader08, csvHeader09, csvHeader10, csvHeader11, csvHeader12, csvHeader13,
                                   csvHeader14, csvHeader15, csvHeader16, csvHeader17, csvHeader18, csvHeader19, csvHeader20,
                                   csvHeader21, csvHeader22, csvHeader23, csvHeader24, csvHeader25, csvHeader26, csvHeader27,
                                   csvHeader28, csvHeader29, csvHeader30 };

            // ヘッダを格納
            for (int i = 0; i < HeadArray.Length; i++)
                dt.Columns.Add(HeadArray[i]);

            int cnt = 0;
            // レコードを格納
            foreach (var item in dataList)
            {
                var row = dt.NewRow();
                foreach (var prop in properties)
                {
                    var itemValue = prop.GetValue(item, new object[] { });
                    row[HeadArray[cnt]] = itemValue;
                    cnt++;
                }
                dt.Rows.Add(row);
                cnt = 0;
            }

            // CSVファイルに書き込む
            using (var writer = new System.IO.StreamWriter(csvPath, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                //ヘッダを書き込む
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    //取得
                    string col = dt.Columns[i].Caption;
                    writer.Write(col);
                    //カンマを書き込む
                    if (dt.Columns.Count - 1 > i)
                    {
                        writer.Write(',');
                    }
                }
                //改行する
                writer.Write("\r\n");

                //レコードを書き込む
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        //取得
                        string col = row[i].ToString();
                        //ダブルクォーテーションで囲む
                        col = EncloseDoubleQuotes(col);
                        //書き込む
                        writer.Write(col);
                        //カンマを書き込む
                        if (dt.Columns.Count - 1 > i)
                        {
                            writer.Write(',');
                        }
                    }
                    writer.Write("\r\n");
                }

                //閉じる
                writer.Close();
            }
        }

        private string EncloseDoubleQuotes(string value)
        {
            return "\"" + value.Replace("\"", "\"\"") + "\""; // ダブルクォーテーションを適切にエスケープ
        }
    }

    public class OutputCsvDaoCombined
    {
        public void OutputCombined(string outputPath)
        {
            // 送り状データを取り込み
            var outputCsvDao = new OutputCsvDao();
            var outputData = new List<PropertyDto>();
            outputCsvDao.Output(outputPath, outputData);

            // 受注一覧データを取り込み
            var outputCsvDaoJ = new OutputCsvDaoJ();
            var outputDataJ = new List<PropertyDtoJ>();
            outputCsvDaoJ.OutputJ(outputPath, outputDataJ);
        }
    }
}