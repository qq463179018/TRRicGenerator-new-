using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Tasks.DataStream
{
    public class DataStreamRicCreationEntity
    {
        public string Ticker { get; set; }
        public string Sedol { get; set; }
        public string CompanyName { get; set; }
        public string FirstTradingDate { get; set; }
        public string ExchangeCode { get; set; }
        public string Isin { get; set; }
        public string SecurityDescription { get; set; }
        public string AssetCategory { get; set; }
        public string SecurityLongDescription { get; set; }
        public string ThomsonReutersClassificationScheme { get; set; }
        public string CUSIP { get; set; }
        public string ReutersEditorialRIC { get; set; }
        public string RIC { get; set; }
        public string CurrencyCode { get; set; }
        public string fileType { get; set; }
        public string roundLotSize { get; set; }
        public string marketSegmentName { get; set; }
        public string originalRecord { get; set; }
        public string fileName { get; set; }

        public DataStreamRicCreationEntity()
        {
            InitializeRicCreation();
        }

        public DataStreamRicCreationEntity(string record)
        {
            if ((record + "").Trim().Length < 1030)
            {
                InitializeRicCreation();
            }
            else
            {
                this.Ticker = GetValue(412, 436, record);
                this.Sedol = GetValue(68, 74, record);
                this.CompanyName = GetValue(105, 184, record);
                this.FirstTradingDate = GetValue(1022, 1029, record);
                this.ExchangeCode = GetValue(98, 100, record);
                this.Isin = GetValue(84, 95, record);
                this.SecurityDescription = GetValue(23, 58, record);
                this.AssetCategory = GetValue(407, 410, record);
                this.SecurityLongDescription = GetValue(437, 546, record);
                this.ThomsonReutersClassificationScheme = GetValue(806, 815, record);
                this.CUSIP = GetValue(59, 67, record);
                this.ReutersEditorialRIC = GetValue(347, 363, record);
                this.RIC = GetValue(3, 22, record);
                this.CurrencyCode = GetValue(101, 103, record);
                this.fileType = string.Empty;
                this.roundLotSize = string.Empty;
                this.marketSegmentName = string.Empty;
                this.originalRecord = string.Empty;
                this.fileName = record.Substring(record.Length - 10, 10);    //EM010612.M
            }
        }

        private void InitializeRicCreation()
        {
            this.Ticker =
            this.Sedol =
            this.CompanyName =
            this.FirstTradingDate =
            this.ExchangeCode =
            this.Isin =
            this.SecurityDescription =
            this.AssetCategory =
            this.SecurityLongDescription =
            this.ThomsonReutersClassificationScheme =
            this.CUSIP =
            this.ReutersEditorialRIC =
            this.RIC =
            this.CurrencyCode =
            this.fileType =
            this.roundLotSize =
            this.marketSegmentName =
            this.originalRecord = string.Empty;
            this.fileName = null;
        }

        private int GetLength(int start, int end)
        {
            if (start >= end)
                return 0;

            return end - start + 1;
        }

        private string GetValue(int start, int end, string record)
        {
            int length = GetLength(start, end);

            if (length <= 0)
                return null;

            return record.Substring(start - 1, length);
        }
    }
}
