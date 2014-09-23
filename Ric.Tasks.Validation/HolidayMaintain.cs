using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Reuters.ProcessQuality.ContentAuto.Lib;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Ric.Db.Manager;
using Ric.Core;
using Ric.Util;


namespace Ric.Tasks.Validation
{
    public enum MarketIdEnum : int
    {
        China = 4,
        HK = 1,
        Japan = 5,
        Korea = 2,
        Thailand = 3,
        TW = 6
    }

    [ConfigStoredInDB]
    public class HolidayMaitainConfig
    {
        public MarketIdEnum MarketId { get; set; }

        [StoreInDB]
        [Description("The query command used to search holiday information in Eikon. E.g. CN/HOLIDAY ")]
        public string QueryCommand { get; set; }
    }

    public class HolidayMaintain : GeneratorBase
    {
        private HolidayMaitainConfig configObj = null;

        protected override void Initialize()
        {
            configObj = Config as HolidayMaitainConfig;
            if (string.IsNullOrEmpty(configObj.QueryCommand))
            {
                string msg = "Query command can't be blank!";
                MessageBox.Show(msg);
                throw new Exception(msg);
            }
        }

        protected override void Start()
        {
            List<HolidayInfo> holiday = GetHolidayInfoFromGats();

            UpdateHolidayToDb(holiday);

            GenerateFile(holiday);
        }

        private void GenerateFile(List<HolidayInfo> holiday)
        {
            string title = "HolidayDate\tMarketId\tMarketName\tComment";
            string[] data = new string[holiday.Count];
            for (int i = 0; i < holiday.Count; i++)
            {
                HolidayInfo item = holiday[i];
                string content = string.Format("{0}\t{1}\t{2}\t{3}", item.HolidayDate, item.MarketId, (MarketIdEnum)item.MarketId, item.Comment);
                data[i] = content;
            }

            string filePath = string.Format("{0}\\{1}_HOLIDAY.txt", GetOutputFilePath(), configObj.MarketId);
            FileUtil.WriteOutputFile(filePath, data, title, WriteMode.Overwrite);
            TaskResultList.Add(new TaskResultEntry("HOLIDAY File","HOLIDAY File",filePath));
        }

        private List<HolidayInfo> GetHolidayInfoFromGats()
        {
            GatsUtil gats = new GatsUtil();
            string response = gats.GetGatsResponse(configObj.QueryCommand.ToUpper(), "");
            if (string.IsNullOrEmpty(response))
            {
                string msg = "Can't get holiday information, please check your command.";
                Logger.Log(msg, Logger.LogType.Error);
                throw new Exception(msg);
            }
            Logger.Log("Get holiday information by GATS. OK!");
            //TaskResultList.Add(new TaskResultEntry("Holiday Infomation", "Holiday Infomation", outputPath));

            List<HolidayInfo> holiday = new List<HolidayInfo>();
            //string[] content = File.ReadAllLines(outputPath);
            string[] content = response.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            string year = string.Empty;
            int startLine = 0;
            for (int i = 0; i < content.Length; i++)
            {
                string row = content[i];
                if (row.Contains("DATES") && row.Contains("MARKET HOLIDAY"))
                {
                    year = Regex.Match(row, @"\d{4}").Value;
                    startLine = i + 1;
                    break;
                }
            }

            for (int i = startLine; i < content.Length; i++)
            {              
                List<string> dateList = FindHolidayInfo(content[i]);//content[i].Substring(42, 6).Trim();
                if (dateList == null || dateList.Count == 0)
                {
                    continue;
                }

                foreach (string dateStr in dateList)
                {
                    string date = dateStr.Replace(" ", "-");
                    string month = date.Split('-')[1];
                    date = date + "-" + year;
                    DateTime dt;
                    if (!DateTime.TryParse(date, out dt))
                    {
                        break;
                    }

                    string comment = content[i].Substring(content[i].IndexOf(month) + month.Length).TrimStart();
                    int blankIndex = comment.IndexOf("   ");
                    if (blankIndex != -1)
                    {
                        comment = comment.Substring(0, blankIndex);
                    }

                    HolidayInfo item = new HolidayInfo();
                    item.HolidayDate = date;
                    item.MarketId = (int)configObj.MarketId;
                    item.Comment = comment;

                    if (item.Comment.ToLower().Contains("half day") && configObj.MarketId.Equals(MarketIdEnum.Japan))
                    {
                        continue;
                    }

                    holiday.Add(item);
                }
            }
            return holiday;
        }

        private List<string> FindHolidayInfo(string line)
        {            
            string pattern = @"(\d{2}|\d{2}-\d{2}) [A-Za-z]+ +";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(line);
            if (!match.Success)
            {
                return null;
            }
            List<string> date = new List<string>();
            string dateMatch = match.Value.Trim();
            if (dateMatch.Contains("-"))
            {
                string[] dayInfo = dateMatch.Split(' ')[0].Split('-');
                string month = dateMatch.Split(' ')[1];
                try
                {
                    int startDate = Convert.ToInt16(dayInfo[0]);
                    int endDate = Convert.ToInt16(dayInfo[1]);
                    for (int i = startDate; i <= endDate; i++)
                    {
                        string datePart = string.Format("{0:00} {1}", i, month);
                        date.Add(datePart);
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format("Error found when try to parse date {0}. Error message: {1}", dateMatch, ex.Message);
                    Logger.Log(msg, Logger.LogType.Error);
                }
            }
            else
            {
                date.Add(dateMatch);
            }
            return date;
        }


        private void UpdateHolidayToDb(List<HolidayInfo> holiday)
        {
            int rows = HolidayManager.UpdateHoliday(holiday);
            string msg = string.Format("Updated {0} holiday record(s) in database. OK!", rows);
            Logger.Log(msg);
        }
    }
}
