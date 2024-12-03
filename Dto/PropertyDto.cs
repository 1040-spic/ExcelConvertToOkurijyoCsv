using System;

namespace ExcelConvertToOkumarukunnCsv.Dto
{
	public class PropertyDto
	{
        public string DealerCd { get; set; }
        public string ExpCd { get; set; }
        public string ExpNm { get; set; }
        public string InvoiceNo { get; set; }
        public string OutDt { get; set; }
        public string OutCd { get; set; }
        public string OutNm { get; set; }
        public string DisAdrs { get; set; }
        public string DisCd { get; set; }
        public string DisNm { get; set; }
        public string ItemCd { get; set; }
        public string JanCd { get; set; }
        public string ItemNm { get; set; }
        public string CallinNm { get; set; }
        public string LotNo { get; set; }
        public string GuideNo { get; set; }
        public string OutQty { get; set; }
        public string Remarks { get; set; }
        public string MatCd { get; set; }
        public string MatNm { get; set; }
        public string OrdDt { get; set; }
        public string DealerOrderNo { get; set; }
        public string HospitalNm { get; set; }
        public string OrderNo { get; set; }
        public string SerialNo { get; set; }
        public string ExpirationDate { get; set; }

        public string OnsetDate { get; set; }
        public string OrderType { get; set; }
        public string DivisionName { get; set; }
        public string Unit { get; set; }
    }

    public class PropertyDtoJ
    {
        public string DealerCd { get; set; }
        public string DataType { get; set; }
        public string OrderNo { get; set; }
        public string OrdDt { get; set; }
        public string InvoiceNo { get; set; }
        public string DisNm { get; set; }
        public string DealerOrderNo { get; set; }
        public string DisAdrs { get; set; }
        public string ItemCd { get; set; }
        public string JanCd { get; set; }
        public string ItemNm { get; set; }
        public string CallinNm { get; set; }
        public string LotNo { get; set; }
        public string SerialNo { get; set; }
        public string ExpirationDate { get; set; }
        public string GuideNo { get; set; }
        public string OrderQty { get; set; }
        public string OutQty { get; set; }
        public string Unit { get; set; }
        public string OutDt { get; set; }
        public string HospitalNm { get; set; }
        public string MatNm { get; set; }
        public string OutNm { get; set; }
        public string ExpCd { get; set; }
        public string ExpNm { get; set; }
        public string Remarks { get; set; }
        public string OnsetDate { get; set; }
        public string OrderType { get; set; }
        public string DivisionName { get; set; }
        public string MailAlert { get; set; }
        public string DisCd { get; set; }
    }
}