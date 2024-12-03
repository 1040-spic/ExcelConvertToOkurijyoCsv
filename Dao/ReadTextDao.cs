using ExcelConvertToOkumarukunnCsv.Common;
using ExcelConvertToOkumarukunnCsv.Dto;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelConvertToOkumarukunnCsv.Dao
{
	class ReadTextDao
	{
        // 配送会社番号の書き換え
        public List<ExpDto> GetCarrierList()
        {
            var fileName = "配送会社設定.txt";
            var encoding = System.Text.Encoding.GetEncoding("SHIFT_JIS");

            var expList = new List<ExpDto>();

            // “配送会社設定”のテキストを読み込む
            using (var reader = new System.IO.StreamReader(fileName, encoding))
            {
                while (!reader.EndOfStream)
                {
                    var record = reader.ReadLine();
                    // 『：』で区切る
                    string[] arr = record.Split(':');

                    // 空白文字を含む行を考慮
                    for (int i = 0; i < arr.Length; i++)
                    {
                        arr[i] = arr[i].Trim();
                    }

                    // 配列の要素数が足りない場合
                    if (arr.Length < 3)
                    {
                        continue; // 現在の行をスキップして次の行へ
                    }

                    // 客先配送会社コード（arr[2]）が空の場合、配送会社コード（arr[1]）と配送会社名（arr[0]）も空に設定
                    string expNm = string.IsNullOrEmpty(arr[2]) ? "" : arr[0]; // ExpNm (配送会社名)
                    string expCd = string.IsNullOrEmpty(arr[2]) ? "" : arr[1]; // ExpCd (配送会社コード)
                    string expKey = string.IsNullOrEmpty(arr[2]) ? "" : arr[2]; // Expkey (客先配送会社コード)

                    var data = new ExpDto()
                    {
                        // 配送会社名[0]:配送会社コード[1]:客先配送会社コード[2]
                        //ExpNm = arr[0],
                        //ExpCd = arr[1],
                        //Expkey = string.IsNullOrEmpty(arr[2]) ? "" : arr[2]
                        ExpNm = expNm, // 配送会社名
                        ExpCd = expCd, // 配送会社コード
                        Expkey = expKey // 客先配送会社コード
                    };
                    expList.Add(data);
                }
            }

            return expList;
            
        }
    }

    class ReadTextDaoJ
    {
        // 配送会社番号の書き換え
        public List<ExpDtoJ> GetCarrierList()
        {
            var fileName = "配送会社設定.txt";
            var encoding = System.Text.Encoding.GetEncoding("SHIFT_JIS");

            var expList = new List<ExpDtoJ>();

            // “配送会社設定”のテキストを読み込む
            using (var reader = new System.IO.StreamReader(fileName, encoding))
            {
                while (!reader.EndOfStream)
                {
                    var record = reader.ReadLine();
                    // 『：』で区切る
                    string[] arr = record.Split(':');

                    // 空白文字を含む行を考慮
                    for (int i = 0; i < arr.Length; i++)
                    {
                        arr[i] = arr[i].Trim();
                    }

                    // 配列の要素数が足りない場合
                    if (arr.Length < 3)
                    {
                        continue; // 現在の行をスキップして次の行へ
                    }

                    // 客先配送会社コード（arr[2]）が空の場合、配送会社コード（arr[1]）と配送会社名（arr[0]）も空に設定
                    string expNm = string.IsNullOrEmpty(arr[2]) ? "" : arr[0]; // ExpNm (配送会社名)
                    string expCd = string.IsNullOrEmpty(arr[2]) ? "" : arr[1]; // ExpCd (配送会社コード)
                    string expKey = string.IsNullOrEmpty(arr[2]) ? "" : arr[2]; // Expkey (客先配送会社コード)

                    var data = new ExpDtoJ()
                    {
                        // 配送会社名[0]:配送会社コード[1]:客先配送会社コード[2]
                        //ExpNm = arr[0],
                        //ExpCd = arr[1],
                        //Expkey = string.IsNullOrEmpty(arr[2]) ? "" : arr[2]
                        ExpNm = expNm, // 配送会社名
                        ExpCd = expCd, // 配送会社コード
                        Expkey = expKey // 客先配送会社コード
                    };

                    expList.Add(data);
                }
            }

            return expList;

        }
    }
}
