using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using Ric.Core;
using Ric.Db.Manager;

namespace Ric.Tasks.China
{
    #region Configuration

    [ConfigStoredInDB]
    public class ChinaIpoFmConfig
    {
        [StoreInDB]
        [Description("The path where the result will be written.\nEg: C:/Mydrive/")]
        public string ResultFolderPath { get; set; }
    }

    #endregion

    #region Task

    class ChinaIpoFm : GeneratorBase
    {
        #region Declaration

        private static ChinaIpoFmConfig configObj = null;
        private static readonly string tableName = "ETI_China_IPO";

        #endregion

        #region Interface implementation

        protected override void Initialize()
        {
            base.Initialize();

            configObj = Config as ChinaIpoFmConfig;
        }

        protected override void Start()
        {
            try
            {
                GenerateFm1();
                GenerateFm2();
            }
            catch (Exception ex)
            {
                throw new Exception("China IPO Fm task failed", ex);
            }
        }

        #endregion

        #region Fm Generation

        private void GenerateFm1()
        {
            string where = String.Format("WHERE Type ='FM1' AND CONVERT(VARCHAR(25), InsertDbDate , 126) LIKE '{0}%'", DateTime.Now.AddDays(-9).ToString("yyyy-MM-dd"));
            var ndaSh = new Nda();
            var ndaSz = new Nda();
            var ndaSz0 = new Nda();
            var ndaSz2 = new Nda();
            var ndaSz3 = new Nda();
            var fm1 = ManagerBase.Select(tableName, new string[] { "*" }, where);
            foreach (DataRow row in fm1.Rows)
            {
                Dictionary<string, string> newInfos = new Dictionary<string, string>
                {
                    {"englishname", row["EnglishName"].ToString()},
                    {"fullname", row["FullName"].ToString()},
                    {"shortname", row["ShortName"].ToString()},
                    {"price", row["Price"].ToString()},
                    {"listingshares", row["ListingShares"].ToString()},
                    {"effectivedate", ((DateTime)row["EffectiveDate"]).ToString("dd-MMM-yyyy")},
                    {"offeringdate", row["OfferingDate"].ToString()},
                    {"insertdbdate", row["InsertDbDate"].ToString()},
                    {"type", row["Type"].ToString()},
                    {"code", row["Code"].ToString()},
                    {"market", row["Market"].ToString()}
                };
                newInfos.Add("traditionalname", ChineseConverter.Convert(newInfos["shortname"], ChineseConversionDirection.SimplifiedToTraditional));
                if (newInfos["market"] == "Shanghai")
                {
                    if (newInfos["code"].EndsWith("5") || newInfos["code"].EndsWith("7") || newInfos["code"].EndsWith("9"))
                    {
                        newInfos.Add("exlname", "SSE_EQB_CNY_ARIC3");
                    }
                    else if (newInfos["code"].EndsWith("4") || newInfos["code"].EndsWith("6") || newInfos["code"].EndsWith("8"))
                    {
                        newInfos.Add("exlname", "SSE_EQB_CNY_EVEN");
                    }
                    else
                    {
                        newInfos.Add("exlname", "SSE_EQB_CNY_ODD");
                    }
                    ndaSh.AddProp(newInfos);
                }
                else
                {
                    if (newInfos["code"].StartsWith("000"))
                    {
                        ndaSz0.AddProp(newInfos);
                    }
                    else if (newInfos["code"].StartsWith("002"))
                    {
                        ndaSz2.AddProp(newInfos);
                    }
                    else
                    {
                        ndaSz3.AddProp(newInfos);
                    }
                    ndaSz.AddProp(newInfos);
                }
            }
            if (ndaSh.format.Prop.Length > 0)
            {
                ndaSh.GenerateAndSave("CnQaAddCNord3", String.Format("{0}SH_QaAddCNord3_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                ndaSh.GenerateAndSave("CnQaChg", String.Format("{0}SH_Chg_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                ndaSh.GenerateAndSave("CnBgChg", String.Format("{0}SH_BgChg_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                ndaSh.GenerateAndSave("CnIdnAddSS", String.Format("{0}SH_IdnAddSS_{1}.txt", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
            if (ndaSz.format.Prop.Length > 0)
            {
                ndaSz.GenerateAndSave("CnQaChg", String.Format("{0}SZ_Chg_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                ndaSz.GenerateAndSave("CnBgChg", String.Format("{0}SZ_BgChg_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                ndaSz.GenerateAndSave("CnIdnAddSZ", String.Format("{0}SZ_IdnAddSZ_{1}.txt", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
            if (ndaSz0.format.Prop.Length > 0)
            {
                ndaSz.GenerateAndSave("CnQaAddCNord4", String.Format("{0}SZ_QaAddCNord4_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
            if (ndaSz2.format.Prop.Length > 0)
            {
                ndaSz.GenerateAndSave("CnQaAddCNord2", String.Format("{0}SZ_QaAddCNord2_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
            if (ndaSz3.format.Prop.Length > 0)
            {
                ndaSz.GenerateAndSave("CnQaAddCNord", String.Format("{0}SZ_QaAddCNord_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
        }

        private void GenerateFm2()
        {
            string where = String.Format("WHERE Type ='FM2' AND CONVERT(VARCHAR(25), InsertDbDate , 126) LIKE '{0}%'", DateTime.Now.AddDays(-9).ToString("yyyy-MM-dd"));
            var nda1 = new Nda();
            var fm2 = ManagerBase.Select(tableName, new string[] { "*" }, where);
            foreach (DataRow row in fm2.Rows)
            {
                Dictionary<string, string> newInfos = new Dictionary<string, string>
                {
                    {"englishname", row["EnglishName"].ToString()},
                    {"fullname", row["FullName"].ToString()},
                    {"shortname", row["ShortName"].ToString()},
                    {"price", row["Price"].ToString()},
                    {"listingshares", row["ListingShares"].ToString()},
                    {"effectivedate", ((DateTime)row["EffectiveDate"]).ToString("dd-MMM-yyyy") },
                    {"offeringdate", row["OfferingDate"].ToString()},
                    {"insertdbdate", row["InsertDbDate"].ToString()},
                    {"type", row["Type"].ToString()},
                    {"code", row["Code"].ToString()},
                    {"market", row["Market"].ToString()}
                };
                newInfos.Add("traditionalname", ChineseConverter.Convert(newInfos["shortname"], ChineseConversionDirection.SimplifiedToTraditional));
                nda1.AddProp(newInfos);
            }
            if (fm2.Rows.Count > 0)
            {
                nda1.GenerateAndSave("CnIaAddFutDat", String.Format("{0}IaAddFutDat_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                nda1.GenerateAndSave("CnLotAdd", String.Format("{0}LotAdd_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                nda1.GenerateAndSave("CnQaAddFutDat", String.Format("{0}QaAddFutDat_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                nda1.GenerateAndSave("CnQaChgFtd", String.Format("{0}QaChgFtd_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
                nda1.GenerateAndSave("CnTickAdd", String.Format("{0}TickAdd_{1}.csv", configObj.ResultFolderPath, DateTime.Now.ToString("ddMM")));
            }
        }
        #endregion
    }

    #endregion
}
