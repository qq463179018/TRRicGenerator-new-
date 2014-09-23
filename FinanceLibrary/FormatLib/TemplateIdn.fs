namespace Ric.FormatLib

module TemplateIdn =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
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

    //-----------------------------------------------------------------------
    //                    Template for TH CW Add
    //-----------------------------------------------------------------------
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

    //-----------------------------------------------------------------------
    // Template for TW CB
    //-----------------------------------------------------------------------

    let TwIdnCb = 
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "EXL_NAME"])
                HLine(["{ric}{type}"; "{displayname} CB{ric}%5,5"; "{ric}{type}"; "{ric}"; "O{ric}"; "****"; "{chinesename}"; "{ric}%4{type}"; "{effectivedateidn}"; "0#{ric}%4rel{type}"; "t{ric}.TWO"; "{ric}"; "{isin}"; "{strike}"; "{ric}"; "OTCTWS_EQLB"])
    ])
