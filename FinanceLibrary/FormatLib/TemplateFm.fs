namespace Ric.FormatLib

module TemplateFm =
    //-----------------------------------------------------------------------
    //                    Template for TH DW Add
    //-----------------------------------------------------------------------
    let ThFm = 
        HFile([
                Titles(["Effective Date"; "SYMBOL"; "DSPLY_NAME"; "RIC"; "OFFCL_CODE"; "EX_SYMBOL"; "BCKGRNDPAG"; "BCAST_REF"; "#INSTMOD_EXPIR_DATE"; "#INSTMOD_LONGLINK1"; "#INSTMOD_LONGLINK2"; "#INSTMOD_MATUR_DATE"; "#INSTMOD_OFFC_CODE2"; "#INSTMOD_STRIKE_PRC"; "#INSTMOD_WNT_RATIO"; "#INSTMOD_MNEMONIC"; "#INSTMOD_TDN_SYMBOL"; "#INSTMOD_LONGLINK3"; "#INSTMOD_GV1_DATE"; "EXL_NAME"; "#INSTMOD_PUTCALLIND"; "Old Chian"; "New BCU"; "NDA Common Name"; "Primary Listing"; "Organisation Name (DIRNAME)"; "Issue Date"; "First Exercise Date"; "Last Exercise Date"; "Last Actual Trading Date"; "Outstanding Warrant Quantity"; "Exercise Period"])
                HLine(["{tradingdate}"; "{ric}.BK"; "{display}"; "{ric}.BK"; "{code}"; "{code}"; "****"; "{asset}.BK"; "{expiredate}"; "{asset}.BK"; "{ric}ta.BK"; "{expiredate}"; ""; "{price}"; "{ratio}"; "{code}"; "{code}"; "0#2{ric}.BK"; "{lasttradingdate}"; "SET_EQLB_W"; "if[{cp}.EQUALS(C)]then[ ]else[PU_PUT]"; "0#IPO.BK"; "0#DW.BK"; 
                "{codestart} {abbr} {bigdate}{cp}WNT"; "{abbr}.BK"; "{name}"; "{tradingdate}"; "{maturitydate}"; "{lastexercisedate}"; "{lasttradingdate}"; "{number}"; "European Style; DW can be exercised only on Automatic Exercise Date."])
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