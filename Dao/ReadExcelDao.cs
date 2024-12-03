using ExcelConvertToOkumarukunnCsv.Common;
using ExcelConvertToOkumarukunnCsv.Dto;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace ExcelConvertToOkumarukunnCsv.Dao
{
    //送り状
	class ReadExcelDao
	{
        // おくまるくん必要項目
        // ディーラーコード・運送会社コード・運送会社名・送り状No・出荷日・発送先コード・出荷先名・品名・品番・ロットNo・
        // 出荷数・備考・担当営業所名・受注日・ディーラー発注番号・病院名・受注No・使用期限・症例日・オーダー種類・単位名

        private enum columnId
		{
            TokuisakiCD,
            OkurijyoNO,
            Syukkabi,
            ButuryuNohinsaki,
            ButuryuNohinsakimei,
            Hinmei,
            Hinmoku,
            RottoNO,
            Suryo,
            Tantosyamei,
            Eigyobumonmei,
            TokuisakiTyumonNO,
            Kokyakumei,
            JyutyuNO,
            SiyoKigen,
            SiyoYoteibi,
            Denpyomei,
            Kashidashikubunmei,
            Tanimei,
            KashidashiKubun,
            HaisoGroup
        }

		public List<ExcelDto> getDataList(string file)
		{
			var excelDataList = new List<ExcelDto>();
			var rowId = 1;

            try
			{
				using (var fs = new FileStream(
						file,
						FileMode.Open,
						FileAccess.Read,
						FileShare.ReadWrite
					)
				)
				{
					var wb = WorkbookFactory.Create(fs);
					var ws = wb.GetSheetAt(0);

					//2行目はヘッダのため、3行目から読み込む
					for (rowId = 2; rowId <= ws.LastRowNum; rowId++)
					{
                        var data = new ExcelDto()
                        {
                            TokuisakiCD = getValue(ws.GetRow(rowId).GetCell(13)),
                            OkurijyoNO = getValue(ws.GetRow(rowId).GetCell(9)),
                            Syukkabi = getValue(ws.GetRow(rowId).GetCell(12)),
                            ButuryuNohinsaki = getValue(ws.GetRow(rowId).GetCell(17)),
                            ButuryuNohinsakimei = getValue(ws.GetRow(rowId).GetCell(18)),
                            Hinmei = getValue(ws.GetRow(rowId).GetCell(24)),
                            Hinmoku = getValue(ws.GetRow(rowId).GetCell(23)),
                            RottoNO = getValue(ws.GetRow(rowId).GetCell(25)),
                            Suryo = getValue(ws.GetRow(rowId).GetCell(28)),
                            Tantosyamei = getValue(ws.GetRow(rowId).GetCell(34)),
                            Eigyobumonmei = getValue(ws.GetRow(rowId).GetCell(32)),
                            TokuisakiTyumonNO = getValue(ws.GetRow(rowId).GetCell(3)),
                            Kokyakumei = getValue(ws.GetRow(rowId).GetCell(16)),
                            JyutyuNO = getValue(ws.GetRow(rowId).GetCell(2)),
                            SiyoKigen = getValue(ws.GetRow(rowId).GetCell(26)),
                            SiyoYoteibi = getValue(ws.GetRow(rowId).GetCell(7)),
                            Denpyomei = getValue(ws.GetRow(rowId).GetCell(5)),
                            Kashidashikubunmei = getValue(ws.GetRow(rowId).GetCell(6)),
                            Tanimei = getValue(ws.GetRow(rowId).GetCell(27)),
                            HaisoGroup = getValue(ws.GetRow(rowId).GetCell(10)),
                        };

                        // Kashidashikubunmeiが「買取」または「長期」の場合、SiyoYoteibiを空欄に設定
                        if (data.Kashidashikubunmei == "買取" || data.Kashidashikubunmei == "長期")
                        {
                            data.SiyoYoteibi = string.Empty; // SiyoYoteibiを空欄に設定
                        }

                        // Denpyomeiが「売上」の場合、Denpyomeiを「買取」に変換
                        if (data.Denpyomei == "売上")
                        {
                            data.Denpyomei = "買取";  // Denpyomeiの値を「買取」に変換
                        }

                        //excelDataList.Add(data);
                        if (!IsEmpty(data))
                        {
                            excelDataList.Add(data);
                        }

                    }
				}
			}
			catch(Exception e)
			{
               MessageBox.Show("Excelの読み込み中にエラーが発生しました。\r\n" + (rowId + 1).ToString() + "行目", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			return excelDataList;
		}

        private bool IsEmpty(ExcelDto data)
        {
            // 必要なフィールドがすべて空かどうかをチェック
            return string.IsNullOrWhiteSpace(data.TokuisakiCD) &&
                   string.IsNullOrWhiteSpace(data.OkurijyoNO) &&
                   string.IsNullOrWhiteSpace(data.Syukkabi) &&
                   string.IsNullOrWhiteSpace(data.ButuryuNohinsaki) &&
                   string.IsNullOrWhiteSpace(data.ButuryuNohinsakimei) &&
                   string.IsNullOrWhiteSpace(data.Hinmei) &&
                   string.IsNullOrWhiteSpace(data.Hinmoku) &&
                   string.IsNullOrWhiteSpace(data.RottoNO) &&
                   string.IsNullOrWhiteSpace(data.Suryo) &&
                   string.IsNullOrWhiteSpace(data.Tantosyamei) &&
                   string.IsNullOrWhiteSpace(data.Eigyobumonmei) &&
                   string.IsNullOrWhiteSpace(data.TokuisakiTyumonNO) &&
                   string.IsNullOrWhiteSpace(data.Kokyakumei) &&
                   string.IsNullOrWhiteSpace(data.JyutyuNO) &&
                   string.IsNullOrWhiteSpace(data.SiyoKigen) &&
                   string.IsNullOrWhiteSpace(data.SiyoYoteibi) &&
                   string.IsNullOrWhiteSpace(data.Denpyomei) &&
                   string.IsNullOrWhiteSpace(data.Kashidashikubunmei) &&
                   string.IsNullOrWhiteSpace(data.Tanimei) &&
                   string.IsNullOrWhiteSpace(data.HaisoGroup);
        }

        private string getValue(ICell cell)
		{
			var str = string.Empty;

			if (cell == null)
				return str;

			switch (cell.CellType)
			{
				case CellType.String:
					//文字列
					str = cell.StringCellValue;
					break;
				case CellType.Numeric:
                    //数値 or 日付
                    if (isCellDateTimeFormatted(cell))
                    {
                        //時刻 or 年月日
                        //ユーザー定義型を考慮して処理
                        if (Constant.CellFormatIndexList.Time.Contains(cell.CellStyle.DataFormat))
                            //時刻
                            str = cell.DateCellValue.ToString("H:mm");
                        else
                            //日付
                            str = cell.DateCellValue.ToString("yyyyMMdd");
                    }

                    else
                        //数値
                        str = cell.NumericCellValue.ToString();
                    break;

                case CellType.Boolean:
					//真偽
					str = cell.BooleanCellValue.ToString();
					break;
				case CellType.Formula:
					//計算式
					str = cell.CellFormula;
					switch (cell.CachedFormulaResultType)
					{
						case CellType.String:
							//文字列
							str = cell.StringCellValue;
							break;
						case CellType.Numeric:
							//数値 or 日付
							if (DateUtil.IsCellDateFormatted(cell))
								//日付
								str = cell.DateCellValue.ToString("yyyyMMdd");
							else
								//数値
								str = cell.NumericCellValue.ToString();
							break;
						case CellType.Boolean:
							//真偽
							str = cell.BooleanCellValue.ToString();
							break;
						default:
							break;
					}
					break;
				default:
					break;
			}

			return str;
		}

		private bool isCellDateTimeFormatted(ICell cell)
		{
			//NPOI標準の日付型判定だとユーザ定義型の場合に日付型として処理されないため
			//プログラム内にユーザ定義型リストを保持
			if (DateUtil.IsCellDateFormatted(cell)
				|| Constant.CellFormatIndexList.Date.Contains(cell.CellStyle.DataFormat)
				|| Constant.CellFormatIndexList.Time.Contains(cell.CellStyle.DataFormat))
				return true;

			return false;
		}
	}

    class ReadExcelDaoJ
    {
        // おくまるくん必要項目

        private enum columnId
        {
            TokuisakiCD,
            JyutyuNO,
            TorokuHiduke,
            ButuryuNohinsakimei,
            TokuisakiTyumonNO,
            Hinmei,
            Hinmoku,
            Suryo,
            Kokyakumei,
            Eigyobumonmei,
            EigyoTantosyamei,
            SiyoYoteibi,
            Denpyomei,
            Kashidashikubunmei,
            JyutyuSeisanKubun,
            KosinNichiji,
            ButuryuNohinsaki,
            HaisoGroup,
            KiboNoki,
            Jyokyo
        }

        public List<ExcelDtoJ> getDataList(string file)
        {
            var excelDataList = new List<ExcelDtoJ>();
            var rowId = 1;
            var today = DateTime.Today;

            try
            {
                using (var fs = new FileStream(
                        file,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite
                    )
                )
                {
                    var wb = WorkbookFactory.Create(fs);
                    var ws = wb.GetSheetAt(0);

                    //3行目はヘッダのため、4行目から読み込む
                    for (rowId = 3; rowId <= ws.LastRowNum; rowId++)
                    {
                        var data = new ExcelDtoJ()
                        {
                            TokuisakiCD = getValue(ws.GetRow(rowId).GetCell(10)),
                            JyutyuNO = getValue(ws.GetRow(rowId).GetCell(2)),
                            TorokuHiduke = getValue(ws.GetRow(rowId).GetCell(47)),
                            ButuryuNohinsakimei = getValue(ws.GetRow(rowId).GetCell(23)),
                            TokuisakiTyumonNO = getValue(ws.GetRow(rowId).GetCell(3)),
                            Hinmei = getValue(ws.GetRow(rowId).GetCell(15)),
                            Hinmoku = getValue(ws.GetRow(rowId).GetCell(14)),
                            Suryo = getValue(ws.GetRow(rowId).GetCell(17)),
                            JikaiNohinYotei = getValue(ws.GetRow(rowId).GetCell(54)),
                            Kokyakumei = getValue(ws.GetRow(rowId).GetCell(13)),
                            Eigyobumonmei = getValue(ws.GetRow(rowId).GetCell(53)),
                            EigyoTantosyamei = getValue(ws.GetRow(rowId).GetCell(51)),
                            SiyoYoteibi = getValue(ws.GetRow(rowId).GetCell(6)),
                            Denpyomei = getValue(ws.GetRow(rowId).GetCell(9)),
                            Kashidashikubunmei = getValue(ws.GetRow(rowId).GetCell(5)),
                            JyutyuSeisanKubun = getValue(ws.GetRow(rowId).GetCell(55)),
                            KosinNichiji = getValue(ws.GetRow(rowId).GetCell(49)),
                            ButuryuNohinsaki = getValue(ws.GetRow(rowId).GetCell(24)),
                            HaisoGroup = getValue(ws.GetRow(rowId).GetCell(26)),
                            KiboNoki = getValue(ws.GetRow(rowId).GetCell(7)),
                            Jyokyo = getValue(ws.GetRow(rowId).GetCell(8))
                        };

                        // JyutyuSeisanKubun の変換処理
                        data.JyutyuSeisanKubun = ConvertJyutyuSeisanKubun(data.JyutyuSeisanKubun);

                        // Kashidashikubunmeiが「買取」または「長期」の場合、SiyoYoteibiを空欄に設定
                        if (data.Kashidashikubunmei == "買取" || data.Kashidashikubunmei == "長期")
                        {
                            // SiyoYoteibiを空欄に設定
                            data.SiyoYoteibi = string.Empty; 
                        }

                        // KiboNokiの値を取得
                        string kiboNoki = data.KiboNoki?.Trim();  // 空白を取り除く
                        if (!string.IsNullOrEmpty(kiboNoki))  // 空欄でない場合のみ処理
                        {
                            DateTime kiboNokiDate;
                            // 期待する日付形式（YYYYMMDD）
                            string[] dateFormats = { "yyyyMMdd" };

                            // KiboNokiが「今日以降」(当日を含まない)かつ、Jyokyoが「後引当以外」の場合、MailAlertに1を設定
                            if (DateTime.TryParseExact(kiboNoki, dateFormats, null, System.Globalization.DateTimeStyles.None, out kiboNokiDate) && kiboNokiDate > today && data.Jyokyo != "後引当")
                            {
                                // 条件に合えば1を設定
                                data.Mail = "1"; 
                            }
                            else
                            {
                                // 条件に合わない場合、空欄に設定
                                data.Mail = string.Empty;
                            }
                        }
                        else
                        {
                            // KiboNokiが空欄の場合は空欄に設定
                            data.Mail = string.Empty;  
                        }
                        //// KiboNokiが「今日以降」、かつ、Jyokyoが「後引当以外」の場合、MailAlertに1を設定
                        //DateTime kiboNokiDate;

                        //if (DateTime.TryParse(data.KiboNoki, out kiboNokiDate) && kiboNokiDate >= today && data.Jyokyo != "後引当")
                        //{
                        //    // 条件に合えば1を設定
                        //    data.Mail = "1";
                        //}
                        //else
                        //{
                        //    // 条件に合わない場合、空欄に設定
                        //    data.Mail = string.Empty;
                        //}

                        //excelDataList.Add(data);
                        if (!IsEmptyData(data))
                        {
                            excelDataList.Add(data);
                        }

                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Excelの読み込み中にエラーが発生しました。\r\n" + (rowId + 1).ToString() + "行目", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return excelDataList;
        }

        // JyutyuSeisanKubun の変換を行う
        private string ConvertJyutyuSeisanKubun(string original)
        {
            // 完全一致の文字列リスト
            string[] exactMatches = new string[]
            {
       　　　　 "ディスコン",
       　　　　 "施設限定品",
        　　　　"海外専用サイズ",
       　　　　 "未上市品",
       　　　　 "受注生産品",
       　　　　 "生産切替確認"
            };

            // 部分一致の文字列リスト
            string[] partialMatches = new string[]
            {
       　　　　 "シラスコン特注品",
       　　　　 "海外代理店専用サイズ"
            };

            // 完全一致の文字列が含まれている場合、そのまま返す
            foreach (var exactMatch in exactMatches)
            {
                if (original == exactMatch)
                {
                    // 完全一致した場合、そのまま返す
                    return original;  
                }
            }

            // 部分一致の文字列が含まれている場合、そのまま返す
            foreach (var partialMatch in partialMatches)
            {
                if (original.Contains(partialMatch))
                {
                    // 部分一致した場合、そのまま返す
                    return original;  
                }
            }

            // それ以外は空欄に設定
            return string.Empty;
           
        }

        private bool IsEmptyData(ExcelDtoJ data)
        {
            return string.IsNullOrEmpty(data.TokuisakiCD) &&
                   string.IsNullOrEmpty(data.JyutyuNO) &&
                   string.IsNullOrEmpty(data.TorokuHiduke) &&
                   string.IsNullOrEmpty(data.ButuryuNohinsakimei) &&
                   string.IsNullOrEmpty(data.TokuisakiTyumonNO) &&
                   string.IsNullOrEmpty(data.Hinmei) &&
                   string.IsNullOrEmpty(data.Hinmoku) &&
                   string.IsNullOrEmpty(data.Suryo) &&
                   string.IsNullOrEmpty(data.JikaiNohinYotei) &&
                   string.IsNullOrEmpty(data.Kokyakumei) &&
                   string.IsNullOrEmpty(data.Eigyobumonmei) &&
                   string.IsNullOrEmpty(data.EigyoTantosyamei) &&
                   string.IsNullOrEmpty(data.SiyoYoteibi) &&
                   string.IsNullOrEmpty(data.Denpyomei) &&
                   string.IsNullOrEmpty(data.Kashidashikubunmei) &&
                   string.IsNullOrEmpty(data.JyutyuSeisanKubun) &&
                   string.IsNullOrEmpty(data.KosinNichiji) &&
                   string.IsNullOrEmpty(data.ButuryuNohinsaki) &&
                   string.IsNullOrEmpty(data.HaisoGroup) &&
                   string.IsNullOrEmpty(data.KiboNoki) &&
                   string.IsNullOrEmpty(data.Jyokyo);
        }

        private string getValue(ICell cell)
        {
            var str = string.Empty;

            if (cell == null)
                return str;

            switch (cell.CellType)
            {
                case CellType.String:
                    //文字列
                    str = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    //数値 or 日付
                    if (isCellDateTimeFormatted(cell))
                    {
                        //時刻 or 年月日
                        //ユーザー定義型を考慮して処理
                        if (Constant.CellFormatIndexList.Time.Contains(cell.CellStyle.DataFormat))
                            //時刻
                            str = cell.DateCellValue.ToString("H:mm");
                        else
                            //日付
                            str = cell.DateCellValue.ToString("yyyyMMdd");
                    }

                    else
                        //数値
                        str = cell.NumericCellValue.ToString();
                    break;

                case CellType.Boolean:
                    //真偽
                    str = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Formula:
                    //計算式
                    str = cell.CellFormula;
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            //文字列
                            str = cell.StringCellValue;
                            break;
                        case CellType.Numeric:
                            //数値 or 日付
                            if (DateUtil.IsCellDateFormatted(cell))
                                //日付
                                str = cell.DateCellValue.ToString("yyyyMMdd");
                            else
                                //数値
                                str = cell.NumericCellValue.ToString();
                            break;
                        case CellType.Boolean:
                            //真偽
                            str = cell.BooleanCellValue.ToString();
                            break;
                        default:
                            break;
                    }
                    break;
                default:
                    break;
            }

            return str;
        }

        private bool isCellDateTimeFormatted(ICell cell)
        {
            //NPOI標準の日付型判定だとユーザ定義型の場合に日付型として処理されないため
            //プログラム内にユーザ定義型リストを保持
            if (DateUtil.IsCellDateFormatted(cell)
                || Constant.CellFormatIndexList.Date.Contains(cell.CellStyle.DataFormat)
                || Constant.CellFormatIndexList.Time.Contains(cell.CellStyle.DataFormat))
                return true;

            return false;
        }
    }
}
