namespace Ric.FormatLib

module Template =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
    let ThFm = 
        HFile([
                Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"; "Old Chian"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name (DIRNAME)"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Last Actual Trading Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
                HLine(["{tradingdate}"; "{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}.BK"; "{expiredate}"; "{asset}.BK"; "{ric}ta.BK"; "{expiredate}"; ""; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdate}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"; "0#IPO.BK"; "0#DW.BK"; 
                "{codestart} {abbr} {bigdate}{cp}WNT"; "{abbr}.BK"; "{name}"; "{tradingdate}"; "{maturitydate}"; "{lastexercisedate}"; "{lasttradingdate}"; "{number}"; "European Style; DW can be exercised only on Automatic Exercise Date."])
    ])

    let QaAdd =
        HFile([
                Titles(["RIC"; "TYPE"; "CATEGORY"; "EXCHANGE"; "CURRENCY"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "EXPIRY DATE"; "STRIKE PRICE"; "TICKER SYMBOL"; "TRADING SEGMENT"; "TAG"; "ROUND LOT SIZE"; "BASE ASSET"; "DERIVATIVES FIRST TRADING DAY"; "CALL PUT OPTION"])
                HLine(["{ric}.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT"; "{display}"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "321"; "100"; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
                HLine(["{ric}ol.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT OL"; "{display}-O"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "47907"; ""; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
    ])

    let DomChain =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "REF_COUNT"; "LINK_1"; "LINK_2"; "EXL_NAME"])
                HLine(["0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "{display}"; "0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "3"; "{ric}ol.BK"; "{ric}ol.BKd"; "SET_EQLB_W_OL_DOM_CHAIN"])
    ])

    let ForIdn =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; 
                        "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; 
                        "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"])
                HLine(["{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}.BK"; "{expiredatewrt}"; "{asset}.BK"; "{ric}ta.BK"; "{expiredatewrt}"; ""; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdatewrt}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"])
    ])

    let IaAdd =
        HFile([
                Titles(["ISIN"; "TYPE"; "CATEGORY"; "WARRANT ISSUER"; "RCS ASSET CLASS"])
                HLine([""; "DERIVATIVE"; "EIW"; "45828"; "TRAD"])
    ])

    let WrtAdd =
        HFile([
                Titles(["Logical_Key"; "Secondary_ID"; "Secondary_ID_Type"; "Warrant_Title"; "Issuer_OrgId"; "Issue_Date"; "Country_Of_Issue"; "Governing_Country"; "Announcement_Date"; "Payment_Date"; "Underlying_Type"; "Clearinghouse1_OrgId"; "Clearinghouse2_OrgId"; "Clearinghouse3_OrgId"; "Guarantor"; "Guarantor_Type"; "Guarantee_Type"; "Incr_Exercise_Lot"; "Min_Exercise_Lot"; "Max_Exercise_Lot"; "Rt_Page_Range"; "Underwriter1_OrgId"; "Underwriter1_Role"; "Underwriter2_OrgId"; "Underwriter2_Role"; "Underwriter3_OrgId"; "Underwriter3_Role"; "Underwriter4_OrgId"; "Underwriter4_Role"; "Exercise_Style"; "Warrant_Type"; "Expiration_Date"; "Registered_Bearer_Code"; "Price_Display_Type"; "Private_Placement"; "Coverage_Type"; "Warrant_Status"; "Status_Date"; "Redemption_Method"; "Issue_Quantity"; "Issue_Price"; "Issue_Currency"; "Issue_Price_Type"; "Issue_Spot_Price"; "Issue_Spot_Currency"; "Issue_Spot_FX_Rate"; "Issue_Delta"; "Issue_Elasticity"; "Issue_Gearing"; "Issue_Premium"; "Issue_Premium_PA"; "Denominated_Amount"; "Exercise_Begin_Date"; "Exercise_End_Date"; "Offset_Number"; "Period_Number"; "Offset_Frequency"; "Offset_Calendar"; "Period_Calendar"; "Period_Frequency"; "RAF_Event_Type"; "Exercise_Price"; "Exercise_Price_Type"; "Warrants_Per_Underlying"; "Underlying_FX_Rate"; "Underlying_RIC"; "Underlying_Item_Quantity"; "Units"; "Cash_Currency"; "Delivery_Type"; "Settlement_Type"; "Settlement_Currency"; "Underlying_Group"; "Country1_Code"; "Coverage1_Type"; "Country2_Code"; "Coverage2_Type"; "Country3_Code"; "Coverage3_Type"; "Country4_Code"; "Coverage4_Type"; "Country5_Code"; "Coverage5_Type"; "Note1_Type"; "Note1"; "Note2_Type"; "Note2"; "Note3_Type"; "Note3"; "Note4_Type"; "Note4"; "Note5_Type"; "Note5"; "Note6_Type"; "Note6"; "Exotic1_Parameter"; "Exotic1_Value"; "Exotic1_Begin_Date"; "Exotic1_End_Date"; "Exotic2_Parameter"; "Exotic2_Value"; "Exotic2_Begin_Date"; "Exotic2_End_Date"; "Exotic3_Parameter"; "Exotic3_Value"; "Exotic3_Begin_Date"; "Exotic3_End_Date"; "Exotic4_Parameter"; "Exotic4_Value"; "Exotic4_Begin_Date"; "Exotic4_End_Date"; "Exotic5_Parameter"; "Exotic5_Value"; "Exotic5_Begin_Date"; "Exotic5_End_Date"; "Exotic6_Parameter"; "Exotic6_Value"; "Exotic6_Begin_Date"; "Exotic6_End_Date"; "Event_Type1"; "Event_Period_Number1"; "Event_Calendar_Type1"; "Event_Frequency1"; "Event_Type2"; "Event_Period_Number2"; "Event_Calendar_Type2"; "Event_Frequency2"; "Exchange_Code1"; "Incr_Trade_Lot1"; "Min_Trade_Lot1"; "Min_Trade_Amount1"; "Exchange_Code2"; "Incr_Trade_Lot2"; "Min_Trade_Lot2"; "Min_Trade_Amount2"; "Exchange_Code3"; "Incr_Trade_Lot3"; "Min_Trade_Lot3"; "Min_Trade_Amount3"; "Exchange_Code4"; "Incr_Trade_Lot4"; "Min_Trade_Lot4"; "Min_Trade_Amount4"; "Attached_To_Id"; "Attached_To_Id_Type"; "Attached_Quantity"; "Attached_Code"; "Detachable_Date"; "Bond_Exercise"; "Bond_Price_Percentage"])
                HLine(["{counter}"; ""; "ISIN"; "{name} SHS if[{cp}.EQUALS(C)]then[CALL WTS {bigexpiredate}]else[PUT WTS {bigexpiredate}]"; ""; "{tradingdatewrt}"; "THA"; "THA"; ""; ""; "STOCK"; ""; ""; ""; ""; ""; ""; "100"; "100"; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; "E"; "if[{cp}.EQUALS(C)]then[Call]else[Put]"; "{expiredatewrt}"; "B"; "D"; ""; "C"; ""; ""; ""; "{rawnumber}"; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; "100"; "{maturitydatewrt}"; "{lastexercisedatewrt}"; ""; ""; ""; ""; ""; ""; ""; "{price}"; "A"; "{ratio}"; ""; "{asset}.BK"; "1"; "shr"; ""; "S"; "S"; "THB"; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; "E"; "European Style; DW can be exercised only on Automatic Exercise Date"; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; ""; "SET"; "100"; "100"])
    ])

    //-----------------------------------------------------------------------
    //                    Template for TH CW Add
    //-----------------------------------------------------------------------
    let WrtAddCw =
        HFile([
                Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"; "Old Chain"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name DIRNAME"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
                HLine(["{tradingdate}"; "{abbr}{warrantnumber}m_t.BK"; "{name}%11-W{warrantnumber}"; "{abbr}{warrantnumber}m_t.BK"; "{symbol}"; "{symbol}"; "****"; "{abbr}m.BK"; "{lastexercisedate}"; "{abbr}m.BK"; "{abbr}{warrantnumber}m_tta.BK"; "{lastexercisedate}"; ""; "{price}"; "{ratio}"; "{symbol}"; "{symbol}"; "0#2{abbr}{warrantnumber}m_t.BK"; "SET_EQLB_W"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {price}CWNT"; ""; "{name}"; "{issuedate}"; "{firstexercisedate}"; "{lastexercisedate}"; "{number}"; ""])
                HLine(["{tradingdate}"; "{abbr}{warrantnumber}m_tn.BK"; "{name}%11-W{warrantnumber}-R"; "{abbr}{warrantnumber}m_tn.BK"; "{symbol}-R"; "{symbol}-R"; "****"; "NA"; "{lastexercisedate}"; "{abbr}{warrantnumber}m_t.BK"; "NA"; "NA"; ""; "{price}"; "{ratio}"; "{symbol}-R"; "{symbol}.R"; "NA"; "SET_EQLB_W_NVDR"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {price}CWNT"; ""; "{name}"; "{issuedate}"; "{firstexercisedate}"; "{lastexercisedate}"; "{number}"; ""])
    ])

    let CwMain =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"])
                HLine(["{abbr}{warrantnumber}_t.BK"; "{abbr}"; "{name}%7-W{warrantnumber}"; "{abbr}{warrantnumber}_t.BK"; "{symbol}"; "{symbol}"; "****"; "{abbr}.BK"; "{lastexercisedate}"; "{abbr}.BK"; "{abbr}{warrantnumber}_tta.BK"; "{lastexercisedate}"; ""; "{price}"; "{ratio}"; "{symbol}"; "{symbol}"; "0#2{abbr}{warrantnumber}_t.BK"; "SET_EQLB_W"])
    ])

    let CwNvdr =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_MKT_SEGMNT"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_WNT_RATIO"; "EXL_NAME"])
                HLine(["{abbr}{warrantnumber}_tn.BK"; "{abbr}"; "{name}%7-W{warrantnumber}-R"; "{abbr}{warrantnumber}_tn.BK"; "{symbol}-R"; "{symbol}-R"; "{lastexercisedate}"; "{abbr}{warrantnumber}_t.BK"; "SET"; "{symbol}-R"; ""; "{price}"; "{symbol}.R"; "{ratio}"; "SET_EQLB_W_NVDR"])
    ])

    //-----------------------------------------------------------------------
    //                    Template for CN FM1
    //-----------------------------------------------------------------------
    let IdnAddSS = 
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "DDS_SYMBOL"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_TDN_ISSUER_NAME"; "EXL_NAME"])
                HLine(["{code}.SS"; ""; "{code}.SS"; "{code}"; "{code}"; "****"; "{traditionalname}"; "{code}.SS"; "{code}"; "t{code}.SS"; "{code}"; ""; "IGNORE"; "{exlname}"])
    ])

    let IdnAddSZ = 
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "DDS_SYMBOL"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_TDN_ISSUER_NAME"; "EXL_NAME"])
                HLine(["{code}.SZ"; ""; "{code}.SZ"; "{code}"; "{code}"; "****"; "{traditionalname}"; "{code}.SZ"; "{code}"; "t{code}.SZ"; "{code}"; ""; "{code}"; "IGNORE"; "SZSE_EQB_CNY_1"])
    ])

    let QaAddCNord =
        HFile([
                Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "TRADING SEGMENT"; "BASE ASSET"; "TAG"])
                HLine(["{code}.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; "{code}"; "100"; "SHZ:CHINEXT"; ""; "179"])
                HLine(["{code}ta.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; "{code}"; ""; ""; ""; "64383"])
    ])

    let QaAddCNord3 =
        HFile([
                Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "BASE ASSET"; "TAG"])
                HLine(["{code}.SS"; ""; ""; "CNY"; "SHH"; "EQUITY"; "ORD"; "{code}"; "100"; ""; "163"])
                HLine(["{code}ta.SS"; ""; ""; "CNY"; "SHH"; "EQUITY"; "ORD"; "{code}"; ""; ""; "64382"])
    ])

    let QaChg =
        HFile([
                Titles(["RIC"; "PRIMARY TRADABLE MARKET QUOTE"])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "Y"])
    ])

    let QaAddCNord2 =
        HFile([
                Titles(["RIC"; "TAG" ; "TICKER SYMBOL"; "BASE ASSET"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "TRADING SEGMENT"; "ROUND LOT SIZE"; "TYPE"; "CATEGORY"; "CURRENCY"; "EXCHANGE"])
                HLine(["{code}.SZ"; "179"; "{code}"; ""; ""; "SHZ:SME"; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
                HLine(["{code}ta.SZ"; "64383"; "{code}"; ""; ""; ""; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
    ])

    let QaAddCNord4 =
        HFile([
                Titles(["RIC"; "TAG" ; "TICKER SYMBOL"; "BASE ASSET"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "TRADING SEGMENT"; "ROUND LOT SIZE"; "TYPE"; "CATEGORY"; "CURRENCY"; "EXCHANGE"])
                HLine(["{code}.SZ"; "179"; "{code}"; ""; ""; "SHZ:XSHE"; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
                HLine(["{code}ta.SZ"; "64383"; "{code}"; ""; ""; ""; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
    ])

    let BgChg =
        HFile([
                Titles(["ORGID"; "PRIMARY ISSUE"; "CHINA SCHEME"])
                HLine([""; ""; " SEE INDUSTRY"])
    ])

    //-----------------------------------------------------------------------
    // Template for CN FM2
    //-----------------------------------------------------------------------
    let IaAddFutDat =
        HFile([
                Titles(["PILC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
                HLine([""; ""; ""; "{effectivedate}"; ""; ""; ""; ""])
                HLine([""; "ISIN"; ""; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine([""; "ASSET COMMON NAME"; "{englishname} Ord Shs A "; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine([""; "RCS ASSET CLASS"; "ORD"; "{effectivedate}"; ""; ""; "PEO"; ""])
    ])

    let LotAdd =
        HFile([
                Titles(["RIC"; "LOT NOT APPLICABLE"; "LOT LADDER NAME"; "LOT EFFECTIVE FROM"; "LOT EFFECTIVE TO"; "LOT PRICE INDICATOR"])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "N"; "LOT_LADDER_EQTY_<100>"; "{effectivedate}"; ""; "ORDER"])
    ])

    let QaAddFutDat =
        HFile([
                Titles(["RIC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; ""; ""; "{effectivedate}"; ""; ""; ""; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "RIC"; "{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "ASSET COMMON NAME"; " ORD A"; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "ROUND LOT SIZE"; "100"; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "TICKER SYMBOL"; "{code}"; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; ""; ""; "{effectivedate}"; ""; ""; ""; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "RIC"; "{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "ASSET COMMON NAME"; " ORD A"; "{effectivedate}"; ""; ""; "PEO"; ""])
                //HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "ROUND LOT SIZE"; ""; "{effectivedate}"; ""; ""; "PEO"; ""])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "TICKER SYMBOL"; "{code}"; "{effectivedate}"; ""; ""; "PEO"; ""])
    ])

    let QaChgFtd =
        HFile([
                Titles(["RIC"; "EQUITY FIRST TRADING DAY"])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[.SS]else[.SZ]"; "{effectivedate}"])
                HLine(["{code}if[{code}.STARTSWITH(6)]then[ta.SS]else[ta.SZ]"; "{effectivedate}"])
    ])

    let TickAdd =
        HFile([
                Titles(["RIC"; "TICK NOT APPLICABLE"; "TICK LADDER NAME"; "TICK EFFECTIVE FROM"; "TICK EFFECTIVE TO"; "TICK PRICE INDICATOR"])
                HLine(["{code}.SS"; "N"; "TICK_LADDER_<0.01>"; "{effectivedate}"; ""; "ORDER"])
    ])

    //-----------------------------------------------------------------------
    // Template for TW CB
    //-----------------------------------------------------------------------

    let TwIdnCb = 
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "EXL_NAME"])
                HLine(["{ric}{type}"; "{displayname} CB{ric}%5,5"; "{ric}{type}"; "{ric}"; "O{ric}"; "****"; "{chinesename}"; "{ric}%4{type}"; "{effectivedateidn}"; "0#{ric}%4rel{type}"; "t{ric}.TWO"; "{ric}"; "{isin}"; "{strike}"; "{ric}"; "OTCTWS_EQLB"])
    ])

    let TwCbBulk = 
        HFile([
                Titles(["Date"; "Display Name"; "RIC"; "ISIN"; "Official Code"])
                HLine(["{effectivedate}"; "{displayname} CB{ric}%5,5"; "{ric}{type}"; "{isin}"; "{ric}"])
    ])

    //-----------------------------------------------------------------------
    //                    Template for TW Future Addition
    //-----------------------------------------------------------------------

    let TwFutureStep11 =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F";"{CODE}F";"{UNDERLYING} {CODE}F FUT";"0#{CODE}Ftm:";"{UNDERLYING}.TW,.{CODE}Fsp.TM"])
                HLine(["TAIFEX_SSF_{CODE}1";"{CODE}1";"{UNDERLYING} {CODE}1 FUT";"0#{CODE}1tm:";"{UNDERLYING}.TW,.{CODE}1sp.TM"])
                HLine(["TAIFEX_SSF_{CODE}2";"{CODE}2";"{UNDERLYING} {CODE}2 FUT";"0#{CODE}2tm:";"{UNDERLYING}.TW,.{CODE}2sp.TM"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD";"{CODE}F";"{UNDERLYING} {CODE}F FUT SPD";"0#{CODE}Ftm-:";"{UNDERLYING}.TW,#NULL#"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD";"{CODE}1";"{UNDERLYING} {CODE}1 FUT SPD";"0#{CODE}1tm-:";"{UNDERLYING}.TW,#NULL#"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD";"{CODE}2";"{UNDERLYING} {CODE}2 FUT SPD";"0#{CODE}2tm-:";"{UNDERLYING}.TW,#NULL#"])
    ])

    let TwFutureStep12 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F"; "MNEMONIC"; "{CODE}F"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD"; "MNEMONIC"; "{CODE}F"])
                HLine(["TAIFEX_SSF_{CODE}1"; "MNEMONIC"; "{CODE}1"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"; "MNEMONIC"; "{CODE}1"])
                HLine(["TAIFEX_SSF_{CODE}2"; "MNEMONIC"; "{CODE}2"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"; "MNEMONIC"; "{CODE}2"])
    ])

    let TwFutureStep13 =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F"; "STQSE"; "\t"; "\t"; "{CODE}Ftm"; "grpcntmarketdatastaff@thomsonreuters.com"])
                HLine(["TAIFEX_SSF_{CODE}1"; "STQSE"; "\t"; "\t"; "{CODE}1tm"; "grpcntmarketdatastaff@thomsonreuters.com"])
                HLine(["TAIFEX_SSF_{CODE}2"; "STQSE"; "\t"; "\t"; "{CODE}2tm"; "grpcntmarketdatastaff@thomsonreuters.com"])   
    ])

    let TwFutureStep14 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD"; "#LEG1_EXL"; "TAIFEX_SSF_{CODE}F"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD"; "#LEG2_EXL"; "TAIFEX_SSF_{CODE}F"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"; "#LEG1_EXL"; "TAIFEX_SSF_{CODE}1"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"; "#LEG2_EXL"; "TAIFEX_SSF_{CODE}1"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"; "#LEG1_EXL"; "TAIFEX_SSF_{CODE}2"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"; "#LEG2_EXL"; "TAIFEX_SSF_{CODE}2"])
    ])

    let TwFutureStep15 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"; "BCAST_REF"; "{UNDERLYING}.TW"])
    ])

    let TwFutureStep16 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F"; "STOCK_RIC"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD"; "STOCK_RIC"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1"; "STOCK_RIC"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"; "STOCK_RIC"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2"; "STOCK_RIC"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"; "STOCK_RIC"; "{UNDERLYING}.TW"])
    ])

    let TwFutureStep17 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD_CHAIN"; "BCAST_REF"; "{UNDERLYING}.TW"])
    ])

    let TwFutureStep18 =
        HFile([
                Titles(["EXL_NAME"; "FIELD_NAME"; "FIELD_VALUE"])
                HLine(["TAIFEX_SSF_{CODE}F_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货價差"])
                HLine(["TAIFEX_SSF_{CODE}1_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货價差"])
                HLine(["TAIFEX_SSF_{CODE}2_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD_CHAIN"; "DSPLY_NMLL"; "{CHINESE NAME}期货價差"])
    ])

    let TwFutureTeSsf =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F"])
                HLine(["TAIFEX_SSF_{CODE}1"])
                HLine(["TAIFEX_SSF_{CODE}2"])
    ])

    let TwFutureTeSpd =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F_SPD"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD"])
    ])

    let TwFutureTeAlias =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F_ALIAS"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD_ALIAS"])
                HLine(["TAIFEX_SSF_{CODE}1_ALIAS"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD_ALIAS"])
                HLine(["TAIFEX_SSF_{CODE}2_ALIAS"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD_ALIAS"])
    ])

    let TwFutureTeChain =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F_CHAIN"])
                HLine(["TAIFEX_SSF_{CODE}F_SPD_CHAIN"])
                HLine(["TAIFEX_SSF_{CODE}1_CHAIN"])
                HLine(["TAIFEX_SSF_{CODE}1_SPD_CHAIN"])
                HLine(["TAIFEX_SSF_{CODE}2_CHAIN"])
                HLine(["TAIFEX_SSF_{CODE}2_SPD_CHAIN"])
    ])

    let TwFutureTeSup =
        HFile([
                HLine(["TAIFEX_SSF_{CODE}F_SUP_CHAIN"])
    ])
