namespace Ric.FormatLib

module Template =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
    let QaAdd =
        HFile([
                Titles(["RIC"; "TYPE"; "CATEGORY"; "EXCHANGE"; "CURRENCY"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "EXPIRY DATE"; "STRIKE PRICE"; "TICKER SYMBOL"; "TRADING SEGMENT"; "TAG"; "ROUND LOT SIZE"; "BASE ASSET"; "DERIVATIVES FIRST TRADING DAY"; "CALL PUT OPTION"])
                HLine(["{ric}.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT"; "{display}"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "321"; "100"; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
                HLine(["{ric}ol.BK"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{codestart} {abbr} {bigdate}{cp}WNT OL"; "{display}-O"; "{expiredate}"; "{price}"; "{code}"; "SET:XBKK"; "47907"; ""; "ISIN:"; "{tradingdate}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
    ])

    let IaAdd =
        HFile([
                Titles(["ISIN"; "TYPE"; "CATEGORY"; "WARRANT ISSUER"; "RCS ASSET CLASS"; "WARRANT ISSUE QUANTITY"])
                HLine([""; "DERIVATIVE"; "EIW"; ""; "TRAD"; "{rawnumber}"])
    ])

    //-----------------------------------------------------------------------
    //                    Template for TH CW Add
    //-----------------------------------------------------------------------
    let QaAddCw =
        HFile([
                Titles(["RIC"; "TYPE"; "CATEGORY"; "EXCHANGE"; "CURRENCY"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "EXPIRY DATE"; "STRIKE PRICE"; "TICKER SYMBOL"; "TRADING SEGMENT"; "TAG"; "ROUND LOT SIZE"; "BASE ASSET"; "DERIVATIVES FIRST TRADING DAY"; "CALL PUT OPTION"])
                HLine(["if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{name}%11 {bigdate}{cp}WNT"; "{name}%11-W{warrantnumber}"; "{lastexercisedateShortYear}"; "{price}"; "{symbol}"; "SET:XBKK"; "321"; "100"; "ISIN:"; "{tradingDateShortYear}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
                HLine(["if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tn.BK]else[{abbr}{warrantnumber}_tn.BK]"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{name}%11 {bigdate}{cp}WNT"; "{name}%11-W{warrantnumber}-R"; "{lastexercisedateShortYear}"; "{price}"; "{symbol}-R"; "SET:XBKK"; "44401"; "100"; "ISIN:"; "{tradingDateShortYear}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
                HLine(["if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tol.BK]else[{abbr}{warrantnumber}_tol.BK]"; "DERIVATIVE"; "EIW"; "SET"; "THB"; "{name}%11 {bigdate}{cp}WNT OL"; "{name}%11-W{warrantnumber}-O"; "{lastexercisedateShortYear}"; "{price}"; "{symbol}"; "SET:XBKK"; "47907"; ""; "ISIN:"; "{tradingDateShortYear}"; "if[{cp}.EQUALS(C)]then[CALL]else[PUT]"])
    ])

    let IaAddCw =
        HFile([
                Titles(["ISIN"; "TYPE"; "CATEGORY"; "WARRANT ISSUER"; "RCS ASSET CLASS"])
                HLine([""; "DERIVATIVE"; "EIW"; ""; "COWNT"; ""])
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

    //-----------------------------------------------------------------------
    //                    Template for TH price change
    //-----------------------------------------------------------------------
    let ThPriceChange =
        HFile([
                Titles(["RIC"; "Symbol"; "ISIN"; "WNT_RATIO"; "WNT_RATIO (NEW)"; "STRIKE_PRC (WT) (OLD)"; "STRIKE_PRC (WT) (NEW)"; "MATUR_DATE"])
                HLine([""; "{symbol}"; ""; "{newratio}"; "{oldratio}"; "{oldprice}"; "{newprice}"; ""])
    ])


    //-----------------------------------------------------------------------
    // Template for TW ORD Add
    //-----------------------------------------------------------------------
    let TwOrdQuoteFutureEmg = 
        HFile([ 
               Titles(["RIC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
               HLine(["{code}.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}.TWO"; "RIC"; "{code}.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}";  ""; ""; "PEO"])
               HLine(["{code}.TWO"; "ROUND LOT SIZE"; "1"; "{effectiveDateLong}";  ""; ""; "PEO"])
               HLine(["{code}.TWO"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}";  ""; ""; "PEO"])
               HLine(["{code}f.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}f.TWO"; "RIC"; "{code}f.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}f.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}f.TWO"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"])
    ])

    let TwOrdQuoteFutureGtsm = 
        HFile([ 
               Titles(["RIC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
               HLine(["{code}.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}.TWO"; "RIC"; "{code}.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TWO"; "ROUND LOT SIZE"; "1000"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TWO"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}stat.TWO"; "RIC"; "{code}stat.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TWO"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}ta.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}ta.TWO"; "RIC"; "{code}ta.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}ta.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}f.TWO"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}f.TWO"; "RIC"; "{code}f.TWO"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}f.TWO"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
    ])

    let TwOrdQuoteFutureTwse = 
        HFile([ 
               Titles(["RIC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
               HLine(["{code}.TW"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}.TW"; "RIC"; "{code}.TW"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TW"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TW"; "ROUND LOT SIZE"; "1000"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}.TW"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TW"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}stat.TW"; "RIC"; "{code}stat.TW"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TW"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}stat.TW"; "TICKER SYMBOL"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}ta.TW"; ""; ""; "{effectiveDateLong}"; ""; ""; ""])
               HLine(["{code}ta.TW"; "RIC"; "{code}ta.TW"; "{effectiveDateLong}"; ""; ""; "PEO"])
               HLine(["{code}ta.TW"; "ASSET COMMON NAME"; "{shortname} ORD"; "{effectiveDateLong}"; ""; ""; "PEO"])
    ])

    let TwOrdIssueFuture =
        HFile([
               Titles(["PILC"; "PROPERTY NAME"; "PROPERTY VALUE"; "EFFECTIVE FROM"; "EFFECTIVE TO"; "CHANGE OFFSET"; "CHANGE TRIGGER"; "CORAX PERMID"])
               HLine([""; ""; ""; "{effectiveDateLong}"; ""; ""; ""; ""])
               HLine([""; "ISIN"; "{isin}"; "{effectiveDateLong}"; ""; ""; "PEO"; ""])
               HLine([""; "TAIWAN CODE"; "{code}"; "{effectiveDateLong}"; ""; ""; "PEO"; ""])
               HLine([""; "ASSET COMMON NAME"; "{displayname} Ord Shs"; "{effectiveDateLong}"; ""; ""; "PEO"; ""])
               HLine([""; "RCS ASSET CLASS"; "ORD"; "{effectiveDateLong}"; ""; ""; "PEO"; ""])
    ])


    let TwOrdAddEmg =
        HFile([
               Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TAG"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "TRADING SEGMENT"; "BASE ASSET"; "EQUITY FIRST TRADING DAY"; "SETTLEMENT PERIOD"])
               HLine(["{code}.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "1127"; "EQUITY"; "ORD"; "{code}"; "1"; "TWO:EMGMKT"; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"])
               HLine(["{code}f.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "67085"; "EQUITY"; "ORD"; "{code}"; ""; ""; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"])
    ])

    let TwOrdAddTwse =
        HFile([
               Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TAG"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "TRADING SEGMENT"; "BASE ASSET"; "EQUITY FIRST TRADING DAY"; "SETTLEMENT PERIOD"; "PRIMARY TRADABLE MARKET QUOTE"])
               HLine(["{code}.TW"; "{shortname} ORD"; "{shortname}"; "TWD"; "TAI"; "126"; "EQUITY"; "ORD"; "{code}"; "1000"; "TAI:XTAI"; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; "Y"])
               HLine(["{code}stat.TW"; "{shortname} ORD"; "{shortname}"; "TWD"; "TAI"; "60006"; "EQUITY"; "ORD"; "{code}"; ""; "TAI:XTAI"; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; ""])
               HLine(["{code}ta.TW"; "{shortname} ORD"; "{shortname}"; "TWD"; "TAI"; "64384"; "EQUITY"; "ORD"; ""; ""; ""; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; ""])
    ])

    let TwOrdAddGtsm =
        HFile([
               Titles(["RIC"; "ASSET COMMON NAME"; "ASSET SHORT NAME"; "CURRENCY"; "EXCHANGE"; "TAG"; "TYPE"; "CATEGORY"; "TICKER SYMBOL"; "ROUND LOT SIZE"; "TRADING SEGMENT"; "BASE ASSET"; "EQUITY FIRST TRADING DAY"; "SETTLEMENT PERIOD"; "PRIMARY TRADABLE MARKET QUOTE"])
               HLine(["{code}.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "645"; "EQUITY"; "ORD"; "{code}"; "1000"; "TWO:ROCO"; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; "Y"])
               HLine(["{code}stat.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "60007"; "EQUITY"; "ORD"; "{code}"; ""; "TWO:ROCO"; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; ""])
               HLine(["{code}ta.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "64385"; "EQUITY"; "ORD"; ""; ""; ""; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; ""])
               HLine(["{code}f.TWO"; "{shortname} ORD"; "{shortname}"; "TWD"; "TWO"; "67085"; "EQUITY"; "ORD"; ""; ""; ""; "ISIN:{isin}"; "{firstTradingDay}"; "T+2"; ""])
    ])

    //-----------------------------------------------------------------------
    // Template for TW ORD Drop
    //-----------------------------------------------------------------------
    let TwOrdDropTwse =
        HFile([
               Titles(["RIC"; "RETIRE DATE"])
               HLine(["{code}.TW"; "{retireDate}"])
               HLine(["{code}stat.TW"; "{retireDate}"])
               HLine(["{code}ta.TW"; "{retireDate}"])
    ])

    let TwOrdDropGtsm =
        HFile([
               Titles(["RIC"; "RETIRE DATE"])
               HLine(["{code}.TWO"; "{retireDate}"])
               HLine(["{code}stat.TWO"; "{retireDate}"])
               HLine(["{code}ta.TWO"; "{retireDate}"])
               HLine(["{code}f.TWO"; "{retireDate}"])
    ])

    let TwOrdDropEmg =
        HFile([
               Titles(["RIC"; "RETIRE DATE"])
               HLine(["{code}.TWO"; "{retireDate}"])
               HLine(["{code}f.TWO"; "{retireDate}"])
               HLine(["{code}ta.TWO"; "{retireDate}"])
    ])