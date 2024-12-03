using ExcelConvertToOkumarukunnCsv.Common;
using ExcelConvertToOkumarukunnCsv.Dao;
using ExcelConvertToOkumarukunnCsv.Dto;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ExcelConvertToOkumarukunnCsv.Services
{
	class FileService
	{
        //CSVファイルを出力する
		public void CreateFile(string file)
		{
            try
            {
                var result = MessageBox.Show("処理を開始します。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);


                // キャンセルボタンが押された場合の処理
                if (result == DialogResult.Cancel)
                {
                    MessageBox.Show("処理がキャンセルされました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var dataList = new ReadExcelDao().getDataList(file);
                new OutputCsvDao().Output(file, convertToCsvDtoList(dataList, file));
                //MessageBox.Show("処理が完了しました。");
                MessageBox.Show("おくまるくん用CSVファイルの作成が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }

        //CSVデータの作成

		private List<PropertyDto> convertToCsvDtoList(List<ExcelDto> excelDataList, string file)
		{
            var carrierList = new ReadTextDao().GetCarrierList();

            var csvDataList = new List<PropertyDto>();
           
              var cnt = 0;
            
            // おくまるくん必要項目をExcelデータと対応させる
			foreach (var excelData in excelDataList)
			{
				cnt++;
				try
				{
                    // おくまるくん必要項目
                    // 出荷先コード・輸送モードNo（運送会社コード）・輸送モード名称（運送会社名）・送り状番号・出荷日・出荷先住所・出荷先名称・発送pcs数
                    // 上記以外は空欄にする
                    var csvData = new PropertyDto()
					{
                        //ディーラーコード（出荷先コード）
                        DealerCd = excelData.TokuisakiCD,
                        //運送会社コード（輸送モードNo）→“配送会社設定.txt”で対応させたもの
                        ExpCd = getCarrierString(carrierList, excelData.HaisoGroup, Constant.Carrier.EmpCd),
                        //運送会社名（輸送モード名称）→“配送会社設定.txt”で対応させたもの
                        ExpNm = getCarrierString(carrierList, excelData.HaisoGroup, Constant.Carrier.EmpNm),
                        //送り状No
                        InvoiceNo = excelData.OkurijyoNO,
                        //出荷日
                        OutDt = excelData.Syukkabi,
                        //出荷元コード
                        OutCd = string.Empty,
                        //出荷元名
                        OutNm = string.Empty,
                        //発送先住所（出荷先住所）
                        DisAdrs = string.Empty,
                        //発送先コード
                        DisCd = excelData.ButuryuNohinsaki,
                        //発送先名（出荷先名称）
                        DisNm = excelData.ButuryuNohinsakimei,
                        //品目コード
                        ItemCd = string.Empty,
                        //JANコード
                        JanCd = string.Empty,
                        //品名
                        ItemNm = excelData.Hinmei,
                        //品番
                        CallinNm = excelData.Hinmoku,
                        //ロットNo
                        LotNo = excelData.RottoNO,
                        //納品書No
                        GuideNo = string.Empty,
                        //出荷数
                        OutQty = excelData.Suryo,
                        //備考
                        Remarks = excelData.Tantosyamei,
                        //担当営業所コード
                        MatCd = string.Empty,
                        //担当営業所名
                        MatNm = excelData.Eigyobumonmei,
                        //受注日
                        OrdDt = excelData.Syukkabi,
                        //ディーラー発注番号
                        DealerOrderNo = excelData.TokuisakiTyumonNO,
                        //病院名
                        HospitalNm = excelData.Kokyakumei,
                        //受注No
                        OrderNo = excelData.JyutyuNO,
                        //シリアルNo
                        SerialNo = string.Empty,
                        //使用期限
                        ExpirationDate = excelData.SiyoKigen,
                        //症例日
                        OnsetDate = excelData.SiyoYoteibi,
                        //オーダー種類　Denpyomeiが「売上」の場合、Denpyomeiを「買取」に変換
                        OrderType = string.IsNullOrEmpty(excelData.Kashidashikubunmei)? excelData.Denpyomei: $"{excelData.Denpyomei} ({excelData.Kashidashikubunmei})",
                        //事業部名
                        DivisionName = string.Empty,
                        //単位名
                        Unit = excelData.Tanimei
                    };
					csvDataList.Add(csvData);
				}
				catch (Exception e)
				{
					MessageBox.Show("CSV作成中にエラーが発生しました。\r\n" + cnt.ToString() + "行目", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
			}

			return csvDataList;
		}

        //運送会社リストを取得
        private string getCarrierString(List<ExpDto> carrierList, string ExpCd, string val)

        {
            string retString = "";
            try
            {
                if (carrierList == null || !carrierList.Any())
                {
                    return retString; // carrierListがnullまたは空の場合、空文字を返す
                }

                if (string.IsNullOrEmpty(ExpCd))
                {
                    return retString;
                }

                var normalizedExpCd = ExpCd.Replace(" ", "");  // 空白を削除

                ExpDto carrier = null;  // ExpDto型に変更

                switch (val)
                {
                    case Constant.Carrier.EmpCd:
                        carrier = carrierList.FirstOrDefault(x => x.Expkey == normalizedExpCd);
                        if (carrier != null)
                        {
                            retString = carrier.ExpCd;
                        }
                        break;

                    case Constant.Carrier.EmpNm:
                        carrier = carrierList.FirstOrDefault(x => x.Expkey == normalizedExpCd);
                        if (carrier != null)
                        {
                            retString = carrier.ExpNm;
                        }
                        break;

                    default:
                        retString = "";
                        break;
                }
            }
            catch (Exception ex)
            {
                retString = "";
            }
            return retString;
        }
    }

    class FileServiceJ
    {
        //CSVファイルを出力する
        public void CreateFile(string file)
        {
            try
            {
                var result = MessageBox.Show("処理を開始します。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);


                // キャンセルボタンが押された場合の処理
                if (result == DialogResult.Cancel)
                {
                    MessageBox.Show("処理がキャンセルされました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var dataList = new ReadExcelDaoJ().getDataList(file);
                new OutputCsvDaoJ().OutputJ(file, convertToCsvDtoListJ(dataList, file));
                MessageBox.Show("おくまるくん用CSVファイルの作成が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }

        //CSVデータの作成

        private List<PropertyDtoJ> convertToCsvDtoListJ(List<ExcelDtoJ> excelDataList, string file)
        {
            var carrierList = new ReadTextDaoJ().GetCarrierList();

            var csvDataList = new List<PropertyDtoJ>();

            var cnt = 0;

            // おくまるくん必要項目をExcelデータと対応させる
            foreach (var excelData in excelDataList)
            {
                cnt++;
                try
                {
                    // オーダー種類を30バイトに収める
                    string orderType = (string.IsNullOrEmpty(excelData.Denpyomei) ? string.Empty :
                                        (excelData.Denpyomei.Trim() == "売上" ? "買取" : excelData.Denpyomei)) +
                                        (string.IsNullOrEmpty(excelData.Kashidashikubunmei) ? string.Empty :
                                         $"({excelData.Kashidashikubunmei})") +
                                        (string.IsNullOrEmpty(excelData.JyutyuSeisanKubun) ? string.Empty :
                                         excelData.JyutyuSeisanKubun);

                    // 30バイト以内に収める処理
                    if (System.Text.Encoding.Default.GetByteCount(orderType) > 30)
                    {
                        orderType = TruncateToByteLength(orderType, 30);
                    }

                    // おくまるくん必要項目以外は空欄にする
                    var csvData = new PropertyDtoJ()
                    {
                        //ディーラーコード（出荷先コード）
                        DealerCd = excelData.TokuisakiCD,
                        //データ区分
                        DataType = "10",
                        //受注No
                        OrderNo = excelData.JyutyuNO,
                        //受注日
                        OrdDt = excelData.TorokuHiduke,
                        //送り状No
                        InvoiceNo = string.Empty,
                        //発送先名（出荷先名称）
                        DisNm = excelData.ButuryuNohinsakimei,
                        //ディーラー発注番号
                        DealerOrderNo = excelData.TokuisakiTyumonNO,
                        //発送先住所（出荷先住所）
                        DisAdrs = string.Empty,
                        //品目コード
                        ItemCd = string.Empty,
                        //JANコード
                        JanCd = string.Empty,
                        //品名
                        ItemNm = excelData.Hinmei,
                        //品番
                        CallinNm = excelData.Hinmoku,
                        //ロットNo
                        LotNo = string.Empty,
                        //シリアルNo
                        SerialNo = string.Empty,
                        //使用期限
                        ExpirationDate = string.Empty,
                        //納品書No
                        GuideNo = string.Empty,
                        //受注数
                        OrderQty = excelData.Suryo,
                        //出荷数
                        OutQty = "0",
                        //単位名
                        Unit = string.Empty,
                        //出荷日
                        OutDt = string.Empty,
                        //病院名
                        HospitalNm = excelData.Kokyakumei,
                        //担当営業所名
                        MatNm = excelData.Eigyobumonmei,
                        //出荷元名
                        OutNm = string.Empty,
                        //運送会社コード（輸送モードNo）→“配送会社設定.txt”で対応させたもの
                        ExpCd = getCarrierString(carrierList, excelData.HaisoGroup, Constant.Carrier.EmpCd),
                        //運送会社名（輸送モード名称）→“配送会社設定.txt”で対応させたもの
                        ExpNm = getCarrierString(carrierList, excelData.HaisoGroup, Constant.Carrier.EmpNm),
                        //備考
                        Remarks = excelData.EigyoTantosyamei,
                        //症例日
                        OnsetDate = excelData.SiyoYoteibi,
                        //オーダー種類　Denpyomei:「売上」/ Denpyomei:「買取」に変換→30バイト内
                        OrderType = orderType,
                        // (string.IsNullOrEmpty(excelData.Denpyomei) ? string.Empty : (excelData.Denpyomei.Trim() == "売上" ? "買取" : excelData.Denpyomei)) +
                        // (string.IsNullOrEmpty(excelData.Kashidashikubunmei) ? string.Empty : $"({excelData.Kashidashikubunmei})") +
                        // (string.IsNullOrEmpty(excelData.JyutyuSeisanKubun) ? string.Empty : excelData.JyutyuSeisanKubun),
                        //事業部名
                        DivisionName = string.Empty,
                        //メール通知区分
                        MailAlert = excelData.Mail,
                        //発送先コード
                        DisCd = excelData.ButuryuNohinsaki

                    };

                    csvDataList.Add(csvData);
                }
                catch (Exception e)
                {
                    MessageBox.Show("CSV作成中にエラーが発生しました。\r\n" + cnt.ToString() + "行目", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

            return csvDataList;
        }

        //運送会社リストを取得
        private string getCarrierString(List<ExpDtoJ> carrierList, string ExpCd, string val)
        {
            string retString = "";
            try
            {
                ExpDtoJ carrier = null;

                switch (val)
                {
                    case Constant.Carrier.EmpCd:
                        carrier = carrierList.FirstOrDefault(x => x.Expkey == ExpCd.Replace(" ", ""));
                        // “配送会社設定.txt”内に該当する客先配送会社番号があるとき→おくまるくん上の配送会社番号を使用
                        if (carrier != null)
                        {
                            retString = carrier.ExpCd;
                        }
                        break;

                    case Constant.Carrier.EmpNm:
                        carrier = carrierList.FirstOrDefault(x => x.Expkey == ExpCd.Replace(" ", ""));
                        // “配送会社設定.txt”内に該当する客先配送会社番号があるとき→おくまるくん上の配送会社名を使用
                        if (carrier != null)
                        {
                            retString = carrier.ExpNm;
                        }
                        break;

                    // “配送会社設定.txt”内に該当する客先配送会社番号がないとき→空文字
                    default:
                        retString = "";
                        break;
                }
            }
            catch (Exception ex)
            {
                retString = "";
            }
            return retString;

        }

        // 文字列を指定したバイト数で切り捨てる
        private string TruncateToByteLength(string input, int maxByteLength)
        {
            int byteCount = System.Text.Encoding.Default.GetByteCount(input);
            if (byteCount <= maxByteLength)
            {
                return input;  // バイト数が指定の長さ以内ならそのまま返す
            }

            // バイト数が指定の長さを超える場合、切り捨てる
            int lengthToKeep = input.Length;
            while (System.Text.Encoding.Default.GetByteCount(input.Substring(0, lengthToKeep)) > maxByteLength)
            {
                lengthToKeep--;
            }

            return input.Substring(0, lengthToKeep);  // 切り捨てた文字列を返す
        }

    }
}
