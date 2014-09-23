using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Tasks.Corax
{
    struct MeetingEventTemplate
    {
        public string RIC { get; set; }
        public string ISIN { get; set; }
        public string MEETUnderlyingOrgId { get; set; }
        public string EventType { get; set; }
        public string EventStatus { get; set; }
        public string MandVollnd { get; set; }
        public string ProcessingStatus { get; set; }
        public string MEETAnnouncementDate { get; set; }
        public string MEETRescindeddate { get; set; }
        public string MEETRecordDate { get; set; }
        public string MeetingStatus { get; set; }
        public string MeetingDate { get; set; }
        public string MeetEndDate1 { get; set; }
        public string MeetingdateTimezonecode { get; set; }
        public string MeetingDate2 { get; set; }
        public string MeetEndDate2 { get; set; }
        public string Meetingdate2Timezonecode { get; set; }
        public string MeetingDate3 { get; set; }
        public string MeetEndDate3 { get; set; }
        public string Meetingdate3Timezonecode { get; set; }
        public string MeetingLocation { get; set; }
        public string LocationUrl { get; set; }
        public string EventUrl { get; set; }
        public string WebcastType { get; set; }
        public string WebcastLink { get; set; }
        public string WebcastExpDate { get; set; }
        public string PriCountrycode { get; set; }
        public string LivePriDial { get; set; }
        public string SecCountrycode { get; set; }
        public string LiveSecDial { get; set; }
        public string LivePassCode { get; set; }
        public string DialInUrl { get; set; }
        public string DialInNotes { get; set; }
        public string ReplayStartDate { get; set; }
        public string ReplayEndDate { get; set; }
        public string ReplayPriDialCountrycode { get; set; }
        public string ReplayPriDial { get; set; }
        public string ReplaySecDialCountrycode { get; set; }
        public string ReplaySecDial { get; set; }
        public string ReplayPassCode { get; set; }
        public string ReplayInNotes { get; set; }
        public string RSVPByDate { get; set; }
        public string RSVPPhoneCountrycode { get; set; }
        public string RSVPPhone { get; set; }
        public string RSVPFaxCountrycode { get; set; }
        public string RSVPFax { get; set; }
        public string RSVPEmail { get; set; }
        public string RSVPUrl { get; set; }
        public string DescriptionType { get; set; }
        public string Description { get; set; }
        public string AnalystNotes { get; set; }
        public string DataEntryStatus { get; set; }
        public string EventSourceESTCode { get; set; }
        public string EventSourceLocalTime { get; set; }
        public string EventSourceTZCode { get; set; }
        public string EventSourceSPCode { get; set; }
        public string EventSourceDescription { get; set; }
        public string EventSourceLink { get; set; }
    }

    struct DivInsertTemplate
    {
        public string URL { get; set; }
        public string AnnualSemiAnnualReport { get; set; }
        public string RIC { get; set; }
        public string DVP_TYPE { get; set; }
        public string CLA_EVENT_STATUS { get; set; }
        public string FPE_PERIOD_END { get; set; }
        public string FPE_PERIOD_LENGTH { get; set; }
        public string DIVIDEND_AMOUNT { get; set; }
        public string ANNOUNCEMENT_DATE { get; set; }
        public string PAY_DATE { get; set; }
        public string RECORD_DATE { get; set; }
        public string EX_DATE { get; set; }
        public string PAYDATE_DAY_SET { get; set; }
        public string PAID_AS_PERCENT { get; set; }
        public string CLA_CUR_VAL { get; set; }
        public string QDI_PERCENT { get; set; }
        public string DESCRIPTION { get; set; }
        public string CLA_MEETING_TYPE_VAL { get; set; }
        public string MEETING_DATE { get; set; }
        public string CLA_TAX_STATUS_VAL { get; set; }
        public string TAX_RATE_PERCENT { get; set; }
        public string FOREIGN_INVESTOR_TAX_RATE { get; set; }
        public string SOURCE_TYPE { get; set; }
        public string RELEASE_DATE { get; set; }
        public string LOCAL_DATE { get; set; }
        public string TIMEZONE_NAME { get; set; }
        public string SOURCE_PROVIDER { get; set; }
        public string BRIDGE_SYMBOL { get; set; }
        public string SEQ_NUM { get; set; }
        public string CLA_RECORD_STATUS { get; set; }
        public string FPE_PERIOD_LENGTH_INDICATOR { get; set; }
        public string CLA_DIV_MARKER_VAL { get; set; }
        public string CLA_DIV_FREQ_VAL { get; set; }
        public string PID_QUARTER { get; set; }
        public string PID_YEAR { get; set; }
        public string PID { get; set; }
        public string CLA_TEXT_TYPE { get; set; }
        public string CLA_DIV_FEATURES_VAL { get; set; }
        public string TAX_CREDIT_PERCENT { get; set; }
        public string CLA_TAX_TRTMNT_MKR_VAL { get; set; }
        public string RESCINDED { get; set; }
        public string CAC_MA_COMMENTS { get; set; }
        public string FRANKED_PERCENT { get; set; }
        public string REINVESTMENT_PLAN_AVAILABLE { get; set; }
        public string REINVESTMENT_DEADLINE { get; set; }
        public string REINVESTMENT_PRICE { get; set; }
        public string CLA_SOURCE_OF_FUND { get; set; }
        public string CLA_DIV_RANKING { get; set; }
        public string DIVIDEND_RANKING_DATE { get; set; }
        public string BOOKCLOSURE_START_DATE { get; set; }
        public string BOOKCLOSURE_END_DATE { get; set; }
        public string MODIFIED { get; set; }
        public string SOURCE_ID { get; set; }
        public string SOURCE_LINK { get; set; }
        public string SOURCE_DESCRIPTION { get; set; }
    }
}
