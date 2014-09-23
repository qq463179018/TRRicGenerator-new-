namespace Ric.FormatLib

module TemplateFm =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
    let ThFm = 
        HFile([
                Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"; "Old Chian"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name (DIRNAME)"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Last Actual Trading Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
                HLine(["{tradingdate}"; "{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}{extension}"; "{expiredate}"; "{asset}{extension}"; "{ric}ta.BK"; "{expiredate}"; ""; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdate}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"; "0#IPO.BK"; "0#DW.BK"; 
                "{codestart} {abbr} {bigdate}{cp}WNT"; "{abbr}.BK"; "{name}"; "{tradingdate}"; "{maturitydate}"; "{lastexercisedate}"; "{lasttradingdate}"; "{number}"; "European Style; DW can be exercised only on Automatic Exercise Date."])
    ])

    //-----------------------------------------------------------------------
    //                    Template for TH CW Add
    //-----------------------------------------------------------------------
    let ThFmCw =
        HFile([
                Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"; "Old Chain"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name DIRNAME"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
                HLine(["{tradingdate}"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "{name}%11-W{warrantnumber}"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "{symbol}"; "{symbol}"; "****"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "{lastexercisedateLongDash}"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tta.BK]else[{abbr}{warrantnumber}_tta.BK]" ; "{lastexercisedateLongWrt}"; ""; "{price}"; "{ratio}"; "{symbol}"; "{symbolDot}"; "if[{market}.EQUALS(mai)]then[0#2{abbr}{warrantnumber}m_t.BK]else[0#2{abbr}{warrantnumber}_t.BK]"; "if[{market}.EQUALS(mai)]then[SET_EQLB_W_MAI]else[SET_EQLB_W]"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {bigdate}CWNT"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "{name} ({abbr})"; "{issuedateWrt}"; "{firstexercisedate}"; "{lastexercisedateLong}"; "{number}"; ""])
                HLine(["{tradingdate}"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tn.BK]else[{abbr}{warrantnumber}_tn.BK]"; "{name}%11-W{warrantnumber}-R"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tn.BK]else[{abbr}{warrantnumber}_tn.BK]"; "{symbol}-R"; "{symbol}-R"; "****"; "NA"; "{lastexercisedateShortYear}"; "{abbr}{warrantnumber}m_t.BK"; "NA"; "NA"; ""; "{price}"; "{ratio}"; "{symbol}-R"; "{symbolDot}.R"; "NA"; "SET_EQLB_W_NVDR"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {bigdate}CWNT"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "{name} ({abbr})"; "{issuedateWrt}"; "{firstexercisedate}"; "{lastexercisedateLong}"; ""; ""])
    ])

    //-----------------------------------------------------------------------
    //                    Template for TH price change
    //-----------------------------------------------------------------------
    let ThPriceChange =
        HFile([
                Titles(["RIC"; "Symbol"; "ISIN"; "WNT_RATIO"; "WNT_RATIO (NEW)"; "STRIKE_PRC (WT) (OLD)"; "STRIKE_PRC (WT) (NEW)"])
                HLine(["{ric}"; "{symbol}"; ""; "{newratio}"; "{oldratio}"; "{oldprice}"; "{newprice}"])
    ])

    //-----------------------------------------------------------------------
    // Template for Taiwan Cb Add
    //-----------------------------------------------------------------------
    let CbAddFm =
        VFile([
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
                      Line("Bond_type", "CONVERTBLE")])
                Footer([Empty
                        Title("=== End of Proforma ===")])
        ])

    //-----------------------------------------------------------------------
    // Template for Taiwan Ord
    //-----------------------------------------------------------------------
    let TwOrdTwse =
        VFile([
                Head([Title("PROFORMA - ADD")
                      Empty
                      NoValueLine("FM Serial Number")])
                Body([Empty
                      Title("For TQS")
                      Line("Effective Date", "{effectivedate}")
                      Line("RIC", "{code}.TWO")
                      Line("Displayname (16)", "{shortname}")
                      Line("Official Code", "{code}")
                      Line("Exchange Symbol", "{code}")
                      Line("OFFC_CODE2", "{isin}")
                      Line("Currency", "TWD")
                      Line("Recordtype", "113")
                      Line("Chain Ric", "")
                      Line("Position in chain", "by alpha order")
                      Line("Lot Size", "1000 Shares")
                      Line("COI DISPLY_NMLL", "")
                      Line("COI SECTOR CHAIN", "")
                      Line("BCAST_REF", "{code}.TW")
                      Separator(50)
                      Empty
                      Title("For MRD")
                      Title("-- NEW ORGANISATION LISTING")
                      Line("Ric", "{code}.TW")
                      Line("ISIN", "{isin}")
                      Line("IDN Longname (36)", "")
                      Line("Issue Classification", "")
                      Line("Primary Listing (Ric)/EDCOID", "")
                      Line("Organisation Name (80)", "")
                      Line("BG Legal Name (Chinese)", "")
                      Line("BG Legal Name", "")
                      Line("Geographical Entity", "")
                      Line("Organisation Type", "")
                      Line("Fund Manager (for Funds)", "")
                      Line("Alias (Previous Name)", "")
                      Line("Alia (General)", "")
                      Line("RBSS Code", "")
                      Line("MSCI Code", "")
                      Line("Business Activity", "")
                      Separator(50)
                      Empty
                      Title("For Maintenance of Local Industrial Sector and Index Constituents and Weightings")
                      Empty
                      Line("RIC", "")
                      Line("Local Sector Classification", "")
                      Line("Index RIC(s)", "")
                      Empty
                      Line("Total Shares Outstanding", "")
                      Separator(50)
                      Empty
                      Title("For TIF")
                      Empty
                      Title("New RICs or Chains")
                      Separator(15)
                      Line("Composite Chain RIC", "0#{code}.TW")
                      Line("Longlink 1 (full Quote RIC)", "{code}.TW")
                      Line("Longlink 2 (TA RIC)", "{code}ta.TW")
                      Line("Longlink 3 (stat RIC)", "{code}stat.TW")
                      Line("Longlink 4 (Commodity Map RIC)", "{code}cmap.TW")
                      Line("Longlink 5 (Stock Relative RICs)", "0#{code}rel.TW")
                      Line("Longlink 6 (TAS RIC)", "t{code}.TW")
                      Line("Longlink 7 (DOM RIC)", "D{code}.TW")
                      Separator(50)
                      Empty
                      Title("Field Mainenance")
                      Separator(15)
                      Title("For Stocks")
                      Line("LONGLINK1", "0#{code}rel.TW")
                      Line("LONGLINK2", "t{code}.TW")
                      Line("Bond_Type", "BOND_TYPE_186")
                      Line("Sector ID", "")
                      Line("SPARE_UBYTE8", "")])
                Footer([Empty
                        Title("=== End of Proforma ===")])
        ])

    let TwOrdGtsm =
        VFile([
                Head([Title("PROFORMA - ADD")
                      Empty
                      NoValueLine("FM Serial Number")])
                Body([Empty
                      Title("TQS")
                      Empty
                      Line("Effective Date", "{effectivedate}")
                      Line("RIC", "{code}.TWO")
                      Line("Displayname (16)", "{shortname}")
                      Line("Official Code", "{code}")
                      Line("Exchange Symbol", "0{code}")
                      Line("OFFC_CODE2", "{isin}")
                      Line("Currency", "TWD")
                      Line("Recordtype", "113")
                      Line("Chain Ric", "")
                      Line("Position in chain", "by alpha order")
                      Line("BCAST_REF", "{code}.TWO")
                      Line("COI DISPLY_NMLL", "{chinesename}")
                      Line("COI SECTOR CHAIN", "")
                      Separator(50)
                      Empty
                      Title("MRD")
                      Line("ISIN", "{isin}")
                      Line("IDN Longname (36)", "{shortname}@STK")
                      Line("Issue Classification", "CS")
                      Line("MSCI Classification", "")
                      Line("Organisation Name (80)", "{fullname}")
                      Line("Organisation Name (Chinese)", "{fullchinesename}")
                      Line("BG Legal Name", "{fullname}")
                      Line("Editorial Primary RIC", "{code}.TWO")
                      Separator(50)
                      Empty
                      Title("Local Sector Maintenance")
                      Separator(30)
                      Line("Local Sector Classification", "")
                      Line("Index RIC(s)", "")
                      Separator(50)
                      Empty
                      Title("For TIF")
                      Empty
                      Title("New RICs or Chains")
                      Line("Composite Chain RIC", "0#{code}.TWO")
                      Line("Longlink 1 (Full Quote RIC)", "{code}.TWO")
                      Line("Longlink 2 (TA RIC)", "{code}ta.TWO")
                      Line("Longlink 3 (stat RIC)", "{code}stat.TWO")
                      Line("Longlink 4 (Commodity Map RIC)", "{code}cmap.TWO")
                      Line("Longlink 5 (Stock Relative RICs)", "0#{code}rel.TWO")
                      Line("Longlink 6 (TAS RIC)", "t{code}.TWO")
                      Line("Longlink 7 (DOM RIC)", "D{code}.TWO")
                      Separator(50)
                      Empty
                      Title("Field Mainenance")
                      Title("For Stocks")
                      Line("LONGLINK1", "0#{code}rel.TWO")
                      Line("LONGLINK2", "t{code}.TWO")
                      Line("Bond_Type", "BOND_TYPE_186")
                      Line("Sector ID", "")
                      Line("SPARE_UBYTE8", "0")])
                Footer([Empty
                        Title("=== End of Proforma ===")])
        ])

    let TwOrdEmg =
        VFile([
                Head([Title("PROFORMA - ADD")
                      Empty
                      NoValueLine("FM Serial Number")])
                Body([Empty
                      Title("For TQS")
                      Line("Effective Date", "{effectivedate}")
                      Line("RIC", "{code}.TWO")
                      Line("Displayname (16)", "{displayname} CB{code}%5,5")
                      Line("Official Code", "")
                      Line("Exchange Symbol", "")
                      Line("OFFC_CODE2", "")
                      Line("Currency", "")
                      Line("Recordtype", "")
                      Line("Chain Ric", "")
                      Line("Position in chain", "")
                      Line("Lot Size", "")
                      Line("COI DISPLY_NMLL", "")
                      Line("COI SECTOR CHAIN", "")
                      Line("BCAST_REF", "")
                      Separator(50)
                      Empty
                      Title("For MRD")
                      Title("-- NEW ORGANISATION LISTING")
                      Line("RIC", "")
                      Line("ISIN", "")
                      Line("IDN Longname (36)", "")
                      Line("Issue Classification", "")
                      Line("Primary Listing (RIC)/EDCOID", "")
                      Line("Organisation Name (80)", "")
                      Line("BG Legal Name (Chinese)", "")
                      Line("BG Legal Name", "")
                      Line("Geographical Entity", "")
                      Line("Organisation Type", "")
                      Line("Fund Manager (for Funds)", "")
                      Line("Alias (Previous Name)", "")
                      Line("Alias (General)", "")
                      Line("RBSS Code", "")
                      Line("MSCI Code", "")
                      Line("Business Activity", "")
                      Separator(50)
                      Empty
                      Title("For TIF")
                      Empty
                      Title("New RICs or Chains")
                      Line("Composite Chain RIC", "")
                      Line("Longlink 1 (full Quote RIC)", "")
                      Line("Longlink 2 (TA RIC)", "")
                      Line("Longlink 3 (stat RIC)", "")
                      Line("Longlink 4 (Commodity Map RIC)", "")
                      Line("Longlink 5 (Stock Relative RICs)", "")
                      Line("Longlink 6 (TAS RIC)", "")
                      Line("Longlink 7 (DOM RIC)", "")
                      Empty
                      Title("Field Mainenance")
                      Empty
                      Title("For Stocks")
                      Line("LONGLINK1", "")
                      Line("LONGLINK2", "")
                      Line("Bond_Type", "")
                      Line("Sector ID", "")
                      Line("SPARE_UBYTE8", "")])
                Footer([Empty
                        Title("=== End of Proforma ===")])
        ])