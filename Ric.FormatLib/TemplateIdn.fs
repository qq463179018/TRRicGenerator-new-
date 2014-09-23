namespace Ric.FormatLib

module TemplateIdn =
    //---------------------------------------------------------------------//
    //                     Template for TH DW Add                          //
    //---------------------------------------------------------------------//
    let DomChain =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "REF_COUNT"; "LINK_1"; "LINK_2"; "EXL_NAME"])
                HLine(["0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "{display}"; "0#{ric}if[{ric}.LENGTH(11)]then[o.BK]else[ol.BK]"; "3"; "{ric}ol.BK"; "{ric}ol.BKd"; "SET_EQLB_W_OL_DOM_CHAIN"])
    ])

    let ForIdn =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; 
                        "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; 
                        "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"; "#INSTMOD_GN_TX20_3"])
                HLine(["{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}{extension}"; "{expiredatewrt}"; "{asset}{extension}"; "{ric}ta.BK"; "{expiredatewrt}"; ""; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdatewrt}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"; "{multiplier}"])
    ])

    //---------------------------------------------------------------------//
    //                     Template for TH CW Add                          //
    //---------------------------------------------------------------------//
    let CwMain =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "EXL_NAME"])
                HLine(["if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "{name}%11-W{warrantnumber}"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "{symbol}"; "{symbol}"; "****"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "{lastexercisedateShort}"; "if[{market}.EQUALS(mai)]then[{abbr}m.BK]else[{abbr}.BK]"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tta.BK]else[{abbr}{warrantnumber}_tta.BK]"; "{lastexercisedateShort}"; ""; "{price}"; "{ratio}"; "{symbol}"; "{symbolDot}"; "if[{market}.EQUALS(mai)]then[0#2{abbr}{warrantnumber}m_t.BK]else[0#2{abbr}{warrantnumber}_t.BK]"; "if[{market}.EQUALS(mai)]then[SET_EQLB_W_MAI]else[SET_EQLB_W]"])
    ])

    let CwNvdr =
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_MKT_SEGMNT"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_WNT_RATIO"; "EXL_NAME"])
                HLine(["if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tn.BK]else[{abbr}{warrantnumber}_tn.BK]"; "{name}%11-W{warrantnumber}-R"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_tn.BK]else[{abbr}{warrantnumber}_tn.BK]"; "{symbol}-R"; "{symbol}-R"; "{lastexercisedate}"; "if[{market}.EQUALS(mai)]then[{abbr}{warrantnumber}m_t.BK]else[{abbr}{warrantnumber}_t.BK]"; "SET"; "{symbol}-R"; ""; "{price}"; "{symbolDot}.R"; "{ratio}"; "SET_EQLB_W_NVDR"])
    ])

    //---------------------------------------------------------------------//
    //                      Template for CN FM1                            //
    //---------------------------------------------------------------------//
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

    //---------------------------------------------------------------------//
    //                       Template for TW CB                            //
    //---------------------------------------------------------------------//

    let TwIdnCb = 
        HFile([
                Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_TDN_SYMBOL"; "EXL_NAME"])
                HLine(["{ric}{type}"; "{displayname} CB{ric}%5,5"; "{ric}{type}"; "{ric}"; "O{ric}"; "****"; "{chinesename}"; "{ric}%4{type}"; "{effectivedateidn}"; "0#{ric}%4rel{type}"; "t{ric}.TWO"; "{ric}"; "{isin}"; "{strike}"; "{ric}"; "OTCTWS_EQLB"])
    ])

    //---------------------------------------------------------------------//
    //                     Template for TW Ord Add                         //
    //---------------------------------------------------------------------//

    let TwOrdIdnTwse =
        HFile([
               Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "X_INST_TITLE"; "#INSTMOD_BKGD_REF"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_TDN_ISSUER_NAME"; "EXL_NAME"])
               HLine(["{code}.TW"; "{shortname}"; "{code}.TW"; "{code}"; "{code}"; "****"; "{chinesename}"; "{code}.TW"; "{shortname}@STK"; "{code}.TWB2"; "0#{code}rel.TW"; "t{code}.TW"; "{code}"; "{isin}"; "{displayname}"; "{code}"; "TAIW_EQB_1"])
    ])

    let TwOrdIdnGtsm =
        HFile([
               Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "#INSTMOD_BKGD_REF"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_TDN_ISSUER_NAME"; "EXL_NAME"])
               HLine(["{code}.TWO"; "{shortname}"; "{code}.TWO"; "{code}"; "{code}"; "****"; "{chinesename}"; "{code}.TWO"; "{code}.TWOB2"; "0#{code}rel.TWO"; "t{code}.TWO"; "{code}"; "{isin}"; "{displayname}"; "{code}"; "OTCTWS_EQB"])
    ])

    let TwOrdIdnEmg =
        HFile([
               Titles(["SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "DSPLY_NMLL"; "BCAST_REF"; "#INSTMOD_BKGD_REF"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MNEMONIC"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_TDN_ISSUER_NAME"; "EXL_NAME"])
               HLine(["{code}.TWO"; "{shortname}"; "{code}.TWO"; "{code}"; "{code}"; "****"; "{chinesename}"; "{code}.TWO"; "{code}.TWOB2"; "0#{code}rel.TWO"; "t{code}.TWO"; "{code}"; "{isin}"; "{displayname}"; "{code}"; "OTCEMG_EQB_1"])
    ])

    //---------------------------------------------------------------------//
    //                     Template for TW ORD Drop                        //
    //---------------------------------------------------------------------//
    let TwOrdDrop =
        HFile([
               Titles(["RIC"])
               HLine(["if[{market}.EQUALS(TWSE)]then[{code}.TW]else[{code}.TWO]"])
    ])