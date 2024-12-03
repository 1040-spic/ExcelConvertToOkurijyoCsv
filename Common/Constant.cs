using System.Collections.Generic;

namespace ExcelConvertToOkumarukunnCsv.Common
{
	class Constant
	{
		public class CellFormatIndexList
		{
			//日付型
			public readonly static IList<int> Date = new List<int> { 31 };
			//31.....yyyy年MM月dd日
			//時刻型
			public readonly static IList<int> Time = new List<int> { 20 };
			//20.....H:mm
		}

		public static readonly Dictionary<string, string> ExpDic = new Dictionary<string, string>()
		{
			//ヤマト
			{ DeliStsUrl.Yamato ,"ヤマト運輸" },
			//日通
			{ DeliStsUrl.Nittsu,"日本通運" },
			//佐川急便
			{ DeliStsUrl.Sagawa ,"佐川急便" },
            //西濃運輸
			{ DeliStsUrl.Seino ,"西濃運輸" },
            //久留米運送
			{ DeliStsUrl.Kurume ,"久留米運送" },
            //札幌通運
			{ DeliStsUrl.Sattsu ,"札幌通運" },
			//近鉄ロジ
			{ DeliStsUrl.Kintetsu ,"近鉄ロジスティクス" },
            //福山通運
            { DeliStsUrl.Fukuyama ,"福山通運" },
            //SBS即配便
            { DeliStsUrl.SBS,"SBS即配便" },
            //日本通運宅配便・アロー便
            { DeliStsUrl.Arrow ,"日本通運宅配便・アロー便" },
            //第一貨物
            { DeliStsUrl.Daiichi ,"第一貨物" },
            //エスラインギフ
            { DeliStsUrl.Sline ,"エスラインギフ" },
            //名鉄運輸
            { DeliStsUrl.Meitetsu ,"名鉄運輸" },
            //トナミ運送(株)
            { DeliStsUrl.Tonami ,"トナミ運送(株)" },
            //岡山県貨物
            { DeliStsUrl.Okayama ,"岡山県貨物" },
            //四国名鉄
            { DeliStsUrl.Shikoku ,"四国名鉄" },
            
        };

		public class DeliStsUrl
		{
			public const string Yamato = "0";
			public const string Nittsu = "1";
			public const string Sagawa = "2";
			public const string Seino = "3";
			public const string Kurume = "4";
			public const string Sattsu = "5";
			public const string Kintetsu = "6";
            public const string Fukuyama = "7";
            public const string SBS = "8";
            public const string Arrow = "9";
            public const string Daiichi = "10";
            public const string Sline = "11";
            public const string Meitetsu = "12";
            public const string Tonami = "13";
            public const string Okayama = "14";
            public const string Shikoku = "15"; 

        }

        public class Carrier
        {
            public const string EmpKey = "empKey";
            public const string EmpCd = "empCd";
            public const string EmpNm = "empNm";
        }
    }
}
