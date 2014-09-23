namespace Finance.Lib

module NdaTemplate =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
    let ThFm =[
            Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"; "Old Chian"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name (DIRNAME)"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Last Actual Trading Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
            HLine(["{tradingdate}"; "{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}.BK"; "{expiredate}"; "{asset}.BK"; "{ric}ta.BK"; "{expiredate}"; " "; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdate}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"; "0#IPO.BK"; "0#DW.BK"; 
            "{codestart} {abbr} {bigdate}{cp}WNT"; " "; "{name}"; "{tradingdate}"; "{maturitydate}"; "{lastexercisedate}"; "{lasttradingdate}"; "\"{number}\""; "European Style; DW can be exercised only on Automatic Exercise Date."])
    ]

    let QaAdd =[
            Titles(["RIC"; "TYPE"; "CATEGORY"; "EXCHANGE"; "CURRENCY"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "EXPIRY DATE"; "STRIKE PRICE"; "TICKER SYMBOL"; "TRADING SEGMENT"; "TAG"; "ROUND LOT SIZE"; "BASE ASSET"; "DERIVATIVES FIRST TRADING DAY"; "CALL PUT OPTION"])
            HLine(["{ric}.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT"; "{display}"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "321"; "100"; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
            HLine(["{ric}ol.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT OL"; "{display}-O"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "47907"; " "; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
    ]

    let DomChain =[
            Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "REF_COUNT"; "LINK_1"; "LINK_2"; "EXL_NAME"])
            HLine(["0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "{display}"; "0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "3"; "{ric}ol.BK"; "{ric}ol.BKd"; "SET_EQLB_W_OL_DOM_CHAIN"])
    ]

    let ForIdn =[
            Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; 
                    "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; 
                    "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"])
            HLine(["{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}.BK"; "{expiredate}"; "{asset}.BK"; "{ric}ta.BK"; "{expiredate}"; " "; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdate}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"])
    ]

    let IaAdd =[
            Titles(["ISIN"; "TYPE"; "CATEGORY"; "WARRANT ISSUER"; "RCS ASSET CLASS"])
            HLine([" "; "DERIVATIVE"; "EIW"; "45828"; "TRAD"])
    ]

    let WrtAdd =[
            Titles(["Logical_Key"; "Secondary_ID"; "Secondary_ID_Type"; "Warrant_Title"; "Issuer_OrgId"; "Issue_Date"; "Country_Of_Issue"; "Governing_Country"; "Announcement_Date"; "Payment_Date"; "Underlying_Type"; "Clearinghouse1_OrgId"; "Clearinghouse2_OrgId"; "Clearinghouse3_OrgId"; "Guarantor"; "Guarantor_Type"; "Guarantee_Type"; "Incr_Exercise_Lot"; "Min_Exercise_Lot"; "Max_Exercise_Lot"; "Rt_Page_Range"; "Underwriter1_OrgId"; "Underwriter1_Role"; "Underwriter2_OrgId"; "Underwriter2_Role"; "Underwriter3_OrgId"; "Underwriter3_Role"; "Underwriter4_OrgId"; "Underwriter4_Role"; "Exercise_Style"; "Warrant_Type"; "Expiration_Date"; "Registered_Bearer_Code"; "Price_Display_Type"; "Private_Placement"; "Coverage_Type"; "Warrant_Status"; "Status_Date"; "Redemption_Method"; "Issue_Quantity"; "Issue_Price"; "Issue_Currency"; "Issue_Price_Type"; "Issue_Spot_Price"; "Issue_Spot_Currency"; "Issue_Spot_FX_Rate"; "Issue_Delta"; "Issue_Elasticity"; "Issue_Gearing"; "Issue_Premium"; "Issue_Premium_PA"; "Denominated_Amount"; "Exercise_Begin_Date"; "Exercise_End_Date"; "Offset_Number"; "Period_Number"; "Offset_Frequency"; "Offset_Calendar"; "Period_Calendar"; "Period_Frequency"; "RAF_Event_Type"; "Exercise_Price"; "Exercise_Price_Type"; "Warrants_Per_Underlying"; "Underlying_FX_Rate"; "Underlying_RIC"; "Underlying_Item_Quantity"; "Units"; "Cash_Currenc"; "Delivery_Type"; "Settlement_Type"; "Settlement_Currency"; "Underlying_Group"; "Country1_Code"; "Coverage1_Type"; "Country2_Code"; "Coverage2_Type"; "Country3_Code"; "Coverage3_Type"; "Country4_Code"; "Coverage4_Type"; "Country5_Code"; "Coverage5_Type"; "Note1_Type"; "Note1,"; "Note2_Type"; "Note2"; "Note3_Type"; "Note3"; "Note4_Type"; "Note4"; "Note5_Type"; "Note5"; "Note6_Type"; "Note6"; "Exotic1_Parameter"; "Exotic1_Value"; "Exotic1_Begin_Date"; "Exotic1_End_Date"; "Exotic2_Parameter"; "Exotic2_Value"; "Exotic2_Begin_Date"; "Exotic2_End_Date"; "Exotic3_Paramete"; "Exotic3_Value"; "Exotic3_Begin_Date"; "Exotic3_End_Date"; "Exotic4_Parameter"; "Exotic4_Value"; "Exotic4_Begin_Date"; "Exotic4_End_Date"; "Exotic5_Parameter"; "Exotic5_Value"; "Exotic5_Begin_Date"; "Exotic5_End_Date"; "Exotic6_Parameter"; "Exotic6_Value"; "Exotic6_Begin_Date"; "Exotic6_End_Date"; "Event_Type1"; "Event_Period_Number1"; "Event_Calendar_Type1"; "Event_Frequency1"; "Event_Type2"; "Event_Period_Number2"; "Event_Calendar_Type2"; "Event_Frequency2"; "Exchange_Code1"; "Incr_Trade_Lot1"; "Min_Trade_Lot1"; "Min_Trade_Amount1"; "Exchange_Code2"; "Incr_Trade_Lot2"; "Min_Trade_Lot2"; "Min_Trade_Amount2"; "Exchange_Code3"; "Incr_Trade_Lot3"; "Min_Trade_Lot3"; "Min_Trade_Amount3"; "Exchange_Code4"; "Incr_Trade_Lot4"; "Min_Trade_Lot4"; "Min_Trade_Amount4"; "Attached_To_Id"; "Attached_To_Id_Type"; "Attached_Quantity"; "Attached_Code"; "Detachable_Date"; "Bond_Exercise"; "Bond_Price_Percentage"])
            HLine(["{counter}"; " "; "ISIN"; "{name} SHS if[{cp}.EQUALS(C)]then[CALL WTS {expiredate}]else[PUT WTS {expiredate}]"; "74826"; "{tradingdatewrt}"; "THA"; "THA"; " "; " "; "STOCK"; " "; " "; " "; " "; " "; " "; "100"; "100"; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; "E"; "if[{cp}.EQUALS(C)]then[Call]else[Put]"; "{expiredatewrt}"; "B"; "D"; " "; "C"; " "; " "; " "; "{rawnumber}"; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; "100"; "{maturitydatewrt}"; "{lastexercisedatewrt}"; " "; " "; " "; " "; " "; " "; " "; "{price}"; "A"; "{ratio}"; " "; "{abbr}.BK"; "1"; "shr"; " "; "S"; "S"; "THB"; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; "E"; "European Style; DW can be exercised only on Automatic Exercise Date"; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; " "; "SET"; "100"; "100"])
    ]

    //-----------------------------------------------------------------------
    //                    Template for TH CW Add
    //-----------------------------------------------------------------------
    let WrtAddCw =[
            Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"; "Old Chain"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name DIRNAME"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
            HLine(["{tradingdate}"; "{abbr}{warrantnumber}m_t.BK"; "{name}%11-W{warrantnumber}"; "{abbr}{warrantnumber}m_t.BK"; "{symbol}"; "{symbol}"; "****"; "{abbr}m.BK"; "{lastexercisedate}"; "{abbr}m.BK"; "{abbr}{warrantnumber}m_tta.BK"; "{lastexercisedate}"; " "; "{price}"; "{ratio}"; "{symbol}"; "{symbol}"; "0#2{abbr}{warrantnumber}m_t.BK"; "SET_EQLB_W"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {price}CWNT"; " "; "{name}"; "{issuedate}"; "{firstexercisedate}"; "{lastexercisedate}"; "{number}"; " "])
            HLine(["{tradingdate}"; "{abbr}{warrantnumber}m_tn.BK"; "{name}%11-W{warrantnumber}-R"; "{abbr}{warrantnumber}m_tn.BK"; "{symbol}-R"; "{symbol}-R"; "****"; "NA"; "{lastexercisedate}"; "{abbr}{warrantnumber}m_t.BK"; "NA"; "NA"; " "; "{price}"; "{ratio}"; "{symbol}-R"; "{symbol}.R"; "NA"; "SET_EQLB_W_NVDR"; "0#IPO.BK"; "0#WRTS-M.BK"; "{name}%11 {price}CWNT"; " "; "{name}"; "{issuedate}"; "{firstexercisedate}"; "{lastexercisedate}"; "{number}"; " "])
    ]

    let Cw =[
            Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"])
            HLine(["{abbr}{warrantnumber}_t.BK"; "{abbr}"; "{name}%7-W{warrantnumber}"; "{abbr}{warrantnumber}_t.BK"; "{symbol}"; "{symbol}"; "****"; "{abbr}.BK"; "{lastexercisedate}"; "{abbr}.BK"; "{abbr}{warrantnumber}_tta.BK"; "{lastexercisedate}"; " "; "{price}"; "{ratio}"; "{symbol}"; "{symbol}"; "0#2{abbr}{warrantnumber}_t.BK"; "SET_EQLB_W"])
    ]

    let CwNvdr =[
            Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_MKT_SEGMNT"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_WNT_RATIO"; "EXL_NAME"])
            HLine(["{abbr}{warrantnumber}_tn.BK"; "{abbr}"; "{name}%7-W{warrantnumber}-R"; "{abbr}{warrantnumber}_tn.BK"; "{symbol}-R"; "{symbol}-R"; "{lastexercisedate}"; "{abbr}{warrantnumber}_t.BK"; "SET"; "{symbol}-R"; " "; "{price}"; "{symbol}.R"; "{ratio}"; "SET_EQLB_W_NVDR"])
    ]

    //-----------------------------------------------------------------------
    //                    Template for CN FM1
    //-----------------------------------------------------------------------
    let QaAddCNord =[
            Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "TRADING SEGMENT"; "BASE ASSET"; "TAG"])
            HLine(["{ric}.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; "{ric}"; "100"; "SHZ:CHINEXT"; "BASEASSET"; "179"])
            HLine(["{ric}ta.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; ""; ""; ""; "BASEASSET"; "64383"])
    ]

    let QaAddCNord3 =[
            Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "BASE ASSET"; "TAG"])
            HLine(["{ric}.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; "{ric}"; "100"; "BASEASSET"; "163"])
            HLine(["{ric}ta.SZ"; ""; ""; "CNY"; "SHZ"; "EQUITY"; "ORD"; ""; ""; "BASEASSET"; "64382"])
    ]

    let QaChg =[
            Titles(["RIC"; "PRIMARY TRADABLE MARKET QUOTE"])
            HLine(["{ric}.SZ"; "Y"])
    ]

    let QaAddCNord2 =[
            Titles(["RIC"; "TAG" ; "TICKER SYMBOL"; "BASE ASSET"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "TRADING SEGMENT"; "ROUND LOT SIZE"; "TYPE"; "CATEGORY"; "CURRENCY"; "EXCHANGE"])
            HLine(["{ric}.SZ"; "179"; "{ric}"; ""; ""; "SHZ:SME"; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
            HLine(["{ric}ta.SZ"; "64383"; ""; ""; ""; ""; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
    ]

    let QaAddCNord4 =[
            Titles(["RIC"; "TAG" ; "TICKER SYMBOL"; "BASE ASSET"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "TRADING SEGMENT"; "ROUND LOT SIZE"; "TYPE"; "CATEGORY"; "CURRENCY"; "EXCHANGE"])
            HLine(["{ric}.SZ"; "179"; "{ric}"; ""; ""; "SHZ:XSHE"; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
            HLine(["{ric}ta.SZ"; "64383"; ""; ""; ""; ""; "100"; "EQUITY"; "ORD"; "CNY"; "SHZ"])
    ]

    let BgChg =[
            Titles(["ORGID"; "PRIMARY ISSUE"; "CHINA SCHEME"])
            HLine([""; ""; "ELC"])
    ]
