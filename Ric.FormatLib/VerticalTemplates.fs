namespace Ric.FormatLib

module FmTemplate = 
    //-----------------------------------------------------------------------
    //                    Template for TW CB Announcement
    //-----------------------------------------------------------------------
    let TwTemplate =[
                    Head([Title("PROFORMA - ADD")
                          Empty
                          NoValueLine("FM Serial Number")])
                    Body([Empty
                          Title("For AQS/TQS")
                          Separator(10)
                          Line("Effective Date", "{effectivedate}")
                          Line("RIC", "{ric}.TWO")
                          Line("Displayname (16)", "{displayname} CB{ric}%5,5")
                          Line("Official Code", "{ric}")
                          Line("Exchange Symbol", "O{ric}")
                          Line("OFFC_CODE2 (ISIN)", "{isin}")
                          Line("Currency", "TWD")
                          Line("Recordtype", "81")
                          Line("Chain Ric", "0#CBND.TWO, 0#{ric}%4rel{type}")
                          Empty
                          Line("Position in chain", "by alpha order")
                          Line("Lot Size", "1000 Shares")
                          Empty
                          Line("COI DISPLY_NMLL", "{chinesename}")
                          Line("COI SECTOR CHAIN", "可轉換公司債 {chinesename}%2")
                          Empty
                          Line("BCAST_REF", "{ric}%4{type}")
                          NoValueLine("WNT_RATIO")
                          Line("STRIKE_PRC (WT, CB)", "{strike}")
                          Line("MATUR_DATE (FI, FL, WT, CB)", "{maturedate}")
                          Line("COUPON RATE (CB)", "0.00%")
                          Separator(40)
                          Empty
                          Title("For MRD")
                          Separator(10)
                          Title("EXISTING ORGANISATION LISTING")
                          Line("RIC", "{ric}.TWO")
                          Line("ISIN", "{isin}")
                          Line("IDN Longname (36)", "{displayname} @0.00 {maturedate}%3,7CNV {ric}%5,5")
                          Line("Issue Classification", "CB")
                          Line("Primary Listing(RIC)/ADCOID", "{ric}%4{type}")
                          Line("Organisation Name (80)", "{name}")
                          Separator(40)
                          Empty
                          Title("For Maintenance of Local Industrial Sector and Index Constituents and Weightings")
                          Empty
                          Line("RIC", "{ric}.TWO")
                          Line("Local Sector Classification", "Convertible Bonds")
                          Line("Index RIC(s)", "n/a")
                          Line("Total Shares Outstanding", "{units} Units")
                          Separator(40)
                          Empty
                          Title("For TIF")
                          Empty
                          Title("New RICs or Chains")
                          Separator(20)
                          Line("Composite Chain RIC", "0#{ric}.TWO")
                          Line("Longlink 1 (Full Quote RIC)", "{ric}.TWO")
                          Line("Longlink 2 (TA RIC)", "{ric}ta.TWO")
                          Line("Longlink 3 (stat RIC)", "{ric}stat.TWO")
                          Line("Longlink 4 (Commodity Map RIC)", "{ric}cmap{type}")
                          Line("Longlink 5 (Stock Relative RICs)", "0#{ric}%4rel{type}")
                          Line("Longlink 6 (TAS RIC)", "t{ric}.TWO")
                          Line("Longlink 7 (DOM RIC)", "D{ric}.TWO")
                          Separator(40)
                          Empty
                          Title("Field Maintenance")
                          Separator(10)
                          Title("For CB")
                          Line("LONGLINK1", "0#{ric}%4rel{type}")
                          Line("LONGLINK2", "t{ric}.TWO")
                          Line("Bond_type", "CONVERTBLE")
                          ])
                    Footer([Empty
                            Title("=== End of Proforma ===")
                           ])
        ]
