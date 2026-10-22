"""Formula snapshot from supplied Actual_data(1).xlsm, I_FPC OA.

Original expressions are retained for audit. Runtime internalizes [17] only.
This captures the supplied 31-day layout, not an arbitrary month adapter.
"""
CLIENT_FORMULAS = {'A1': "='[17]TOTAL FGT OA'!A4",
 'A13': '=+WEEKDAY(B13)',
 'B13': '=TEXT(\'[17]TOTAL FGT OA\'!B12,"mm/dd/yyyy")',
 'F13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M13': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B13)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B13)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N13': '=SUM(F13:M13)',
 'O13': '=(N13+C13)',
 'P13': '=IF(D13="","NA",(C13*D13))',
 'Q13': '=(O13+Q12)',
 'S13': '=IF(I13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T13': '=IF(M13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U13': '=IF(H13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V13': '=IF(G13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W13': '=IF(K13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X13': '=IF(L13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y13': '=IF(J13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z13': '=IF(F13=0,"",ROUND(VLOOKUP(DATEVALUE(B13)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA13': '=B13',
 'AD13': '=I13+K13+L13',
 'AE13': "='[17]GSNG Imbalance'!Y11",
 'AF13': '=AD13+AE13',
 'A14': '=+WEEKDAY(B14)',
 'B14': '=TEXT(\'[17]TOTAL FGT OA\'!B13,"mm/dd/yyyy")',
 'F14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M14': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B14)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B14)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N14': '=SUM(F14:M14)',
 'O14': '=(N14+C14)',
 'P14': '=IF(D14="","NA",(C14*D14))',
 'Q14': '=(Q13+O14)',
 'S14': '=IF(I14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T14': '=IF(M14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U14': '=IF(H14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V14': '=IF(G14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W14': '=IF(K14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X14': '=IF(L14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y14': '=IF(J14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z14': '=IF(F14=0,"",ROUND(VLOOKUP(DATEVALUE(B14)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA14': '=B14',
 'AD14': '=I14+K14+L14',
 'AE14': "='[17]GSNG Imbalance'!Y12",
 'AF14': '=AD14+AE14',
 'A15': '=+WEEKDAY(B15)',
 'B15': '=TEXT(\'[17]TOTAL FGT OA\'!B14,"mm/dd/yyyy")',
 'F15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M15': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B15)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B15)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N15': '=SUM(F15:M15)',
 'O15': '=(N15+C15)',
 'P15': '=IF(D15="","NA",(C15*D15))',
 'Q15': '=(Q14+O15)',
 'S15': '=IF(I15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T15': '=IF(M15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U15': '=IF(H15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V15': '=IF(G15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W15': '=IF(K15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X15': '=IF(L15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y15': '=IF(J15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z15': '=IF(F15=0,"",ROUND(VLOOKUP(DATEVALUE(B15)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA15': '=B15',
 'AD15': '=I15+K15+L15',
 'AE15': "='[17]GSNG Imbalance'!Y13",
 'AF15': '=AD15+AE15',
 'A16': '=+WEEKDAY(B16)',
 'B16': '=TEXT(\'[17]TOTAL FGT OA\'!B15,"mm/dd/yyyy")',
 'F16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M16': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B16)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B16)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N16': '=SUM(F16:M16)',
 'O16': '=(N16+C16)',
 'P16': '=IF(D16="","NA",(C16*D16))',
 'Q16': '=(Q15+O16)',
 'S16': '=IF(I16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T16': '=IF(M16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U16': '=IF(H16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V16': '=IF(G16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W16': '=IF(K16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X16': '=IF(L16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y16': '=IF(J16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z16': '=IF(F16=0,"",ROUND(VLOOKUP(DATEVALUE(B16)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA16': '=B16',
 'AD16': '=I16+K16+L16',
 'AE16': "='[17]GSNG Imbalance'!Y14",
 'AF16': '=AD16+AE16',
 'A17': '=+WEEKDAY(B17)',
 'B17': '=TEXT(\'[17]TOTAL FGT OA\'!B16,"mm/dd/yyyy")',
 'F17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M17': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B17)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B17)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N17': '=SUM(F17:M17)',
 'O17': '=(N17+C17)',
 'P17': '=IF(D17="","NA",(C17*D17))',
 'Q17': '=(Q16+O17)',
 'S17': '=IF(I17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T17': '=IF(M17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U17': '=IF(H17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V17': '=IF(G17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W17': '=IF(K17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X17': '=IF(L17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y17': '=IF(J17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z17': '=IF(F17=0,"",ROUND(VLOOKUP(DATEVALUE(B17)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA17': '=B17',
 'AD17': '=I17+K17+L17',
 'AE17': "='[17]GSNG Imbalance'!Y15",
 'AF17': '=AD17+AE17',
 'A18': '=+WEEKDAY(B18)',
 'B18': '=TEXT(\'[17]TOTAL FGT OA\'!B17,"mm/dd/yyyy")',
 'F18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M18': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B18)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B18)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N18': '=SUM(F18:M18)',
 'O18': '=(N18+C18)',
 'P18': '=IF(D18="","NA",(C18*D18))',
 'Q18': '=(Q17+O18)',
 'S18': '=IF(I18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T18': '=IF(M18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U18': '=IF(H18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V18': '=IF(G18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W18': '=IF(K18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X18': '=IF(L18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y18': '=IF(J18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z18': '=IF(F18=0,"",ROUND(VLOOKUP(DATEVALUE(B18)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA18': '=B18',
 'AD18': '=I18+K18+L18',
 'AE18': "='[17]GSNG Imbalance'!Y16",
 'AF18': '=AD18+AE18',
 'A19': '=+WEEKDAY(B19)',
 'B19': '=TEXT(\'[17]TOTAL FGT OA\'!B18,"mm/dd/yyyy")',
 'F19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M19': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B19)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B19)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N19': '=SUM(F19:M19)',
 'O19': '=(N19+C19)',
 'P19': '=IF(D19="","NA",(C19*D19))',
 'Q19': '=(Q18+O19)',
 'S19': '=IF(I19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T19': '=IF(M19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U19': '=IF(H19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V19': '=IF(G19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W19': '=IF(K19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X19': '=IF(L19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y19': '=IF(J19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z19': '=IF(F19=0,"",ROUND(VLOOKUP(DATEVALUE(B19)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA19': '=B19',
 'AD19': '=I19+K19+L19',
 'AE19': "='[17]GSNG Imbalance'!Y17",
 'AF19': '=AD19+AE19',
 'A20': '=+WEEKDAY(B20)',
 'B20': '=TEXT(\'[17]TOTAL FGT OA\'!B19,"mm/dd/yyyy")',
 'F20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M20': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B20)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B20)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N20': '=SUM(F20:M20)',
 'O20': '=(N20+C20)',
 'P20': '=IF(D20="","NA",(C20*D20))',
 'Q20': '=(Q19+O20)',
 'S20': '=IF(I20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T20': '=IF(M20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U20': '=IF(H20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V20': '=IF(G20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W20': '=IF(K20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X20': '=IF(L20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y20': '=IF(J20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z20': '=IF(F20=0,"",ROUND(VLOOKUP(DATEVALUE(B20)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA20': '=B20',
 'AD20': '=I20+K20+L20',
 'AE20': "='[17]GSNG Imbalance'!Y18",
 'AF20': '=AD20+AE20',
 'A21': '=+WEEKDAY(B21)',
 'B21': '=TEXT(\'[17]TOTAL FGT OA\'!B20,"mm/dd/yyyy")',
 'F21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M21': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B21)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B21)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N21': '=SUM(F21:M21)',
 'O21': '=(N21+C21)',
 'P21': '=IF(D21="","NA",(C21*D21))',
 'Q21': '=(Q20+O21)',
 'S21': '=IF(I21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T21': '=IF(M21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U21': '=IF(H21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V21': '=IF(G21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W21': '=IF(K21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X21': '=IF(L21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y21': '=IF(J21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z21': '=IF(F21=0,"",ROUND(VLOOKUP(DATEVALUE(B21)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA21': '=B21',
 'AD21': '=I21+K21+L21',
 'AE21': "='[17]GSNG Imbalance'!Y19",
 'AF21': '=AD21+AE21',
 'A22': '=+WEEKDAY(B22)',
 'B22': '=TEXT(\'[17]TOTAL FGT OA\'!B21,"mm/dd/yyyy")',
 'F22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M22': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B22)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B22)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N22': '=SUM(F22:M22)',
 'O22': '=(N22+C22)',
 'P22': '=IF(D22="","NA",(C22*D22))',
 'Q22': '=(Q21+O22)',
 'S22': '=IF(I22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T22': '=IF(M22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U22': '=IF(H22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V22': '=IF(G22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W22': '=IF(K22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X22': '=IF(L22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y22': '=IF(J22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z22': '=IF(F22=0,"",ROUND(VLOOKUP(DATEVALUE(B22)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA22': '=B22',
 'AD22': '=I22+K22+L22',
 'AE22': "='[17]GSNG Imbalance'!Y20",
 'AF22': '=AD22+AE22',
 'A23': '=+WEEKDAY(B23)',
 'B23': '=TEXT(\'[17]TOTAL FGT OA\'!B22,"mm/dd/yyyy")',
 'F23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M23': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B23)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B23)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N23': '=SUM(F23:M23)',
 'O23': '=(N23+C23)',
 'P23': '=IF(D23="","NA",(C23*D23))',
 'Q23': '=(Q22+O23)',
 'S23': '=IF(I23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T23': '=IF(M23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U23': '=IF(H23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V23': '=IF(G23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W23': '=IF(K23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X23': '=IF(L23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y23': '=IF(J23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z23': '=IF(F23=0,"",ROUND(VLOOKUP(DATEVALUE(B23)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA23': '=B23',
 'AD23': '=I23+K23+L23',
 'AE23': "='[17]GSNG Imbalance'!Y21",
 'AF23': '=AD23+AE23',
 'A24': '=+WEEKDAY(B24)',
 'B24': '=TEXT(\'[17]TOTAL FGT OA\'!B23,"mm/dd/yyyy")',
 'F24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M24': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B24)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B24)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N24': '=SUM(F24:M24)',
 'O24': '=(N24+C24)',
 'P24': '=IF(D24="","NA",(C24*D24))',
 'Q24': '=(Q23+O24)',
 'S24': '=IF(I24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T24': '=IF(M24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U24': '=IF(H24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V24': '=IF(G24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W24': '=IF(K24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X24': '=IF(L24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y24': '=IF(J24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z24': '=IF(F24=0,"",ROUND(VLOOKUP(DATEVALUE(B24)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA24': '=B24',
 'AD24': '=I24+K24+L24',
 'AE24': "='[17]GSNG Imbalance'!Y22",
 'AF24': '=AD24+AE24',
 'A25': '=+WEEKDAY(B25)',
 'B25': '=TEXT(\'[17]TOTAL FGT OA\'!B24,"mm/dd/yyyy")',
 'F25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M25': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B25)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B25)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N25': '=SUM(F25:M25)',
 'O25': '=(N25+C25)',
 'P25': '=IF(D25="","NA",(C25*D25))',
 'Q25': '=(Q24+O25)',
 'S25': '=IF(I25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T25': '=IF(M25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U25': '=IF(H25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V25': '=IF(G25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W25': '=IF(K25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X25': '=IF(L25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y25': '=IF(J25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z25': '=IF(F25=0,"",ROUND(VLOOKUP(DATEVALUE(B25)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA25': '=B25',
 'AD25': '=I25+K25+L25',
 'AE25': "='[17]GSNG Imbalance'!Y23",
 'AF25': '=AD25+AE25',
 'A26': '=+WEEKDAY(B26)',
 'B26': '=TEXT(\'[17]TOTAL FGT OA\'!B25,"mm/dd/yyyy")',
 'F26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M26': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B26)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B26)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N26': '=SUM(F26:M26)',
 'O26': '=(N26+C26)',
 'P26': '=IF(D26="","NA",(C26*D26))',
 'Q26': '=(Q25+O26)',
 'S26': '=IF(I26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T26': '=IF(M26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U26': '=IF(H26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V26': '=IF(G26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W26': '=IF(K26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X26': '=IF(L26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y26': '=IF(J26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z26': '=IF(F26=0,"",ROUND(VLOOKUP(DATEVALUE(B26)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA26': '=B26',
 'AD26': '=I26+K26+L26',
 'AE26': "='[17]GSNG Imbalance'!Y24",
 'AF26': '=AD26+AE26',
 'A27': '=+WEEKDAY(B27)',
 'B27': '=TEXT(\'[17]TOTAL FGT OA\'!B26,"mm/dd/yyyy")',
 'F27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M27': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B27)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B27)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N27': '=SUM(F27:M27)',
 'O27': '=(N27+C27)',
 'P27': '=IF(D27="","NA",(C27*D27))',
 'Q27': '=(Q26+O27)',
 'S27': '=IF(I27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T27': '=IF(M27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U27': '=IF(H27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V27': '=IF(G27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W27': '=IF(K27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X27': '=IF(L27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y27': '=IF(J27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z27': '=IF(F27=0,"",ROUND(VLOOKUP(DATEVALUE(B27)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA27': '=B27',
 'AD27': '=I27+K27+L27',
 'AE27': "='[17]GSNG Imbalance'!Y25",
 'AF27': '=AD27+AE27',
 'A28': '=+WEEKDAY(B28)',
 'B28': '=TEXT(\'[17]TOTAL FGT OA\'!B27,"mm/dd/yyyy")',
 'F28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M28': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B28)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B28)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N28': '=SUM(F28:M28)',
 'O28': '=(N28+C28)',
 'P28': '=IF(D28="","NA",(C28*D28))',
 'Q28': '=(Q27+O28)',
 'S28': '=IF(I28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T28': '=IF(M28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U28': '=IF(H28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V28': '=IF(G28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W28': '=IF(K28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X28': '=IF(L28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y28': '=IF(J28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z28': '=IF(F28=0,"",ROUND(VLOOKUP(DATEVALUE(B28)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA28': '=B28',
 'AD28': '=I28+K28+L28',
 'AE28': "='[17]GSNG Imbalance'!Y26",
 'AF28': '=AD28+AE28',
 'A29': '=+WEEKDAY(B29)',
 'B29': '=TEXT(\'[17]TOTAL FGT OA\'!B28,"mm/dd/yyyy")',
 'F29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M29': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B29)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B29)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N29': '=SUM(F29:M29)',
 'O29': '=(N29+C29)',
 'P29': '=IF(D29="","NA",(C29*D29))',
 'Q29': '=(Q28+O29)',
 'S29': '=IF(I29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T29': '=IF(M29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U29': '=IF(H29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V29': '=IF(G29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W29': '=IF(K29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X29': '=IF(L29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y29': '=IF(J29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z29': '=IF(F29=0,"",ROUND(VLOOKUP(DATEVALUE(B29)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA29': '=B29',
 'AD29': '=I29+K29+L29',
 'AE29': "='[17]GSNG Imbalance'!Y27",
 'AF29': '=AD29+AE29',
 'A30': '=+WEEKDAY(B30)',
 'B30': '=TEXT(\'[17]TOTAL FGT OA\'!B29,"mm/dd/yyyy")',
 'F30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M30': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B30)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B30)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N30': '=SUM(F30:M30)',
 'O30': '=(N30+C30)',
 'P30': '=IF(D30="","NA",(C30*D30))',
 'Q30': '=(Q29+O30)',
 'S30': '=IF(I30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T30': '=IF(M30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U30': '=IF(H30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V30': '=IF(G30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W30': '=IF(K30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X30': '=IF(L30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y30': '=IF(J30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z30': '=IF(F30=0,"",ROUND(VLOOKUP(DATEVALUE(B30)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA30': '=B30',
 'AD30': '=I30+K30+L30',
 'AE30': "='[17]GSNG Imbalance'!Y28",
 'AF30': '=AD30+AE30',
 'A31': '=+WEEKDAY(B31)',
 'B31': '=TEXT(\'[17]TOTAL FGT OA\'!B30,"mm/dd/yyyy")',
 'F31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M31': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B31)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B31)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N31': '=SUM(F31:M31)',
 'O31': '=(N31+C31)',
 'P31': '=IF(D31="","NA",(C31*D31))',
 'Q31': '=(Q30+O31)',
 'S31': '=IF(I31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T31': '=IF(M31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U31': '=IF(H31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V31': '=IF(G31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W31': '=IF(K31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X31': '=IF(L31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y31': '=IF(J31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z31': '=IF(F31=0,"",ROUND(VLOOKUP(DATEVALUE(B31)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA31': '=B31',
 'AD31': '=I31+K31+L31',
 'AE31': "='[17]GSNG Imbalance'!Y29",
 'AF31': '=AD31+AE31',
 'A32': '=+WEEKDAY(B32)',
 'B32': '=TEXT(\'[17]TOTAL FGT OA\'!B31,"mm/dd/yyyy")',
 'F32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M32': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B32)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B32)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N32': '=SUM(F32:M32)',
 'O32': '=(N32+C32)',
 'P32': '=IF(D32="","NA",(C32*D32))',
 'Q32': '=(Q31+O32)',
 'S32': '=IF(I32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T32': '=IF(M32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U32': '=IF(H32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V32': '=IF(G32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W32': '=IF(K32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X32': '=IF(L32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y32': '=IF(J32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z32': '=IF(F32=0,"",ROUND(VLOOKUP(DATEVALUE(B32)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA32': '=B32',
 'AD32': '=I32+K32+L32',
 'AE32': "='[17]GSNG Imbalance'!Y30",
 'AF32': '=AD32+AE32',
 'A33': '=+WEEKDAY(B33)',
 'B33': '=TEXT(\'[17]TOTAL FGT OA\'!B32,"mm/dd/yyyy")',
 'F33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M33': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B33)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B33)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N33': '=SUM(F33:M33)',
 'O33': '=(N33+C33)',
 'P33': '=IF(D33="","NA",(C33*D33))',
 'Q33': '=(Q32+O33)',
 'S33': '=IF(I33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T33': '=IF(M33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U33': '=IF(H33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V33': '=IF(G33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W33': '=IF(K33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X33': '=IF(L33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y33': '=IF(J33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z33': '=IF(F33=0,"",ROUND(VLOOKUP(DATEVALUE(B33)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA33': '=B33',
 'AD33': '=I33+K33+L33',
 'AE33': "='[17]GSNG Imbalance'!Y31",
 'AF33': '=AD33+AE33',
 'A34': '=+WEEKDAY(B34)',
 'B34': '=TEXT(\'[17]TOTAL FGT OA\'!B33,"mm/dd/yyyy")',
 'F34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M34': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B34)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B34)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N34': '=SUM(F34:M34)',
 'O34': '=(N34+C34)',
 'P34': '=IF(D34="","NA",(C34*D34))',
 'Q34': '=(Q33+O34)',
 'S34': '=IF(I34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T34': '=IF(M34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U34': '=IF(H34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V34': '=IF(G34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W34': '=IF(K34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X34': '=IF(L34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y34': '=IF(J34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z34': '=IF(F34=0,"",ROUND(VLOOKUP(DATEVALUE(B34)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA34': '=B34',
 'AD34': '=I34+K34+L34',
 'AE34': "='[17]GSNG Imbalance'!Y32",
 'AF34': '=AD34+AE34',
 'A35': '=+WEEKDAY(B35)',
 'B35': '=TEXT(\'[17]TOTAL FGT OA\'!B34,"mm/dd/yyyy")',
 'F35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M35': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B35)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B35)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N35': '=SUM(F35:M35)',
 'O35': '=(N35+C35)',
 'P35': '=IF(D35="","NA",(C35*D35))',
 'Q35': '=(Q34+O35)',
 'S35': '=IF(I35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T35': '=IF(M35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U35': '=IF(H35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V35': '=IF(G35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W35': '=IF(K35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X35': '=IF(L35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y35': '=IF(J35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z35': '=IF(F35=0,"",ROUND(VLOOKUP(DATEVALUE(B35)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA35': '=B35',
 'AD35': '=I35+K35+L35',
 'AE35': "='[17]GSNG Imbalance'!Y33",
 'AF35': '=AD35+AE35',
 'A36': '=+WEEKDAY(B36)',
 'B36': '=TEXT(\'[17]TOTAL FGT OA\'!B35,"mm/dd/yyyy")',
 'F36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M36': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B36)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B36)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N36': '=SUM(F36:M36)',
 'O36': '=(N36+C36)',
 'P36': '=IF(D36="","NA",(C36*D36))',
 'Q36': '=(Q35+O36)',
 'S36': '=IF(I36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T36': '=IF(M36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U36': '=IF(H36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V36': '=IF(G36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W36': '=IF(K36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X36': '=IF(L36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y36': '=IF(J36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z36': '=IF(F36=0,"",ROUND(VLOOKUP(DATEVALUE(B36)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA36': '=B36',
 'AD36': '=I36+K36+L36',
 'AE36': "='[17]GSNG Imbalance'!Y34",
 'AF36': '=AD36+AE36',
 'A37': '=+WEEKDAY(B37)',
 'B37': '=TEXT(\'[17]TOTAL FGT OA\'!B36,"mm/dd/yyyy")',
 'F37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M37': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B37)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B37)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N37': '=SUM(F37:M37)',
 'O37': '=(N37+C37)',
 'P37': '=IF(D37="","NA",(C37*D37))',
 'Q37': '=(Q36+O37)',
 'S37': '=IF(I37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T37': '=IF(M37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U37': '=IF(H37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V37': '=IF(G37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W37': '=IF(K37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X37': '=IF(L37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y37': '=IF(J37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z37': '=IF(F37=0,"",ROUND(VLOOKUP(DATEVALUE(B37)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA37': '=B37',
 'AD37': '=I37+K37+L37',
 'AE37': "='[17]GSNG Imbalance'!Y35",
 'AF37': '=AD37+AE37',
 'A38': '=+WEEKDAY(B38)',
 'B38': '=TEXT(\'[17]TOTAL FGT OA\'!B37,"mm/dd/yyyy")',
 'F38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M38': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B38)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B38)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N38': '=SUM(F38:M38)',
 'O38': '=(N38+C38)',
 'P38': '=IF(D38="","NA",(C38*D38))',
 'Q38': '=(Q37+O38)',
 'S38': '=IF(I38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T38': '=IF(M38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U38': '=IF(H38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V38': '=IF(G38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W38': '=IF(K38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X38': '=IF(L38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y38': '=IF(J38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z38': '=IF(F38=0,"",ROUND(VLOOKUP(DATEVALUE(B38)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA38': '=B38',
 'AD38': '=I38+K38+L38',
 'AE38': "='[17]GSNG Imbalance'!Y36",
 'AF38': '=AD38+AE38',
 'A39': '=+WEEKDAY(B39)',
 'B39': '=TEXT(\'[17]TOTAL FGT OA\'!B38,"mm/dd/yyyy")',
 'F39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M39': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B39)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B39)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N39': '=SUM(F39:M39)',
 'O39': '=(N39+C39)',
 'P39': '=IF(D39="","NA",(C39*D39))',
 'Q39': '=(Q38+O39)',
 'S39': '=IF(I39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T39': '=IF(M39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U39': '=IF(H39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V39': '=IF(G39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W39': '=IF(K39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X39': '=IF(L39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y39': '=IF(J39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z39': '=IF(F39=0,"",ROUND(VLOOKUP(DATEVALUE(B39)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA39': '=B39',
 'AD39': '=I39+K39+L39',
 'AE39': "='[17]GSNG Imbalance'!Y37",
 'AF39': '=AD39+AE39',
 'A40': '=+WEEKDAY(B40)',
 'B40': '=TEXT(\'[17]TOTAL FGT OA\'!B39,"mm/dd/yyyy")',
 'F40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M40': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B40)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B40)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N40': '=SUM(F40:M40)',
 'O40': '=(N40+C40)',
 'P40': '=IF(D40="","NA",(C40*D40))',
 'Q40': '=(Q39+O40)',
 'S40': '=IF(I40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T40': '=IF(M40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U40': '=IF(H40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V40': '=IF(G40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W40': '=IF(K40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X40': '=IF(L40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y40': '=IF(J40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z40': '=IF(F40=0,"",ROUND(VLOOKUP(DATEVALUE(B40)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA40': '=B40',
 'AD40': '=I40+K40+L40',
 'AE40': "='[17]GSNG Imbalance'!Y38",
 'AF40': '=AD40+AE40',
 'A41': '=+WEEKDAY(B41)',
 'B41': '=TEXT(\'[17]TOTAL FGT OA\'!B40,"mm/dd/yyyy")',
 'F41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M41': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B41)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B41)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N41': '=SUM(F41:M41)',
 'O41': '=(N41+C41)',
 'P41': '=IF(D41="","NA",(C41*D41))',
 'Q41': '=(Q40+O41)',
 'S41': '=IF(I41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T41': '=IF(M41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U41': '=IF(H41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V41': '=IF(G41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W41': '=IF(K41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X41': '=IF(L41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y41': '=IF(J41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z41': '=IF(F41=0,"",ROUND(VLOOKUP(DATEVALUE(B41)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA41': '=B41',
 'AD41': '=I41+K41+L41',
 'AE41': "='[17]GSNG Imbalance'!Y39",
 'AF41': '=AD41+AE41',
 'A42': '=+WEEKDAY(B42)',
 'B42': '=TEXT(\'[17]TOTAL FGT OA\'!B41,"mm/dd/yyyy")',
 'F42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M42': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B42)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B42)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N42': '=SUM(F42:M42)',
 'O42': '=(N42+C42)',
 'P42': '=IF(D42="","NA",(C42*D42))',
 'Q42': '=(Q41+O42)',
 'S42': '=IF(I42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T42': '=IF(M42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U42': '=IF(H42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V42': '=IF(G42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W42': '=IF(K42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X42': '=IF(L42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y42': '=IF(J42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z42': '=IF(F42=0,"",ROUND(VLOOKUP(DATEVALUE(B42)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA42': '=B42',
 'AD42': '=I42+K42+L42',
 'AE42': "='[17]GSNG Imbalance'!Y40",
 'AF42': '=AD42+AE42',
 'A43': '=+WEEKDAY(B43)',
 'B43': '=TEXT(\'[17]TOTAL FGT OA\'!B42,"mm/dd/yyyy")',
 'F43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$F$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'G43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$G$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'H43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$H$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'I43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$I$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'J43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$J$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'K43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$K$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'L43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$L$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'M43': "=IF(ISNA(VALUE(VLOOKUP(DATEVALUE(B43)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE))),0,VALUE(VLOOKUP(DATEVALUE(B43)&$M$6,'[17]FGT Data Download - "
        "New'!$A$4:$R$1048576,14,FALSE)))",
 'N43': '=SUM(F43:M43)',
 'O43': '=(N43+C43)',
 'P43': '=IF(D43="","NA",(C43*D43))',
 'Q43': '=(Q42+O43)',
 'S43': '=IF(I43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$I$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'T43': '=IF(M43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$M$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'U43': '=IF(H43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$H$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'V43': '=IF(G43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$G$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'W43': '=IF(K43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$K$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'X43': '=IF(L43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$L$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Y43': '=IF(J43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$J$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'Z43': '=IF(F43=0,"",ROUND(VLOOKUP(DATEVALUE(B43)&$F$6,\'[17]FGT Data Download - '
        "New'!$A:$R,17,FALSE)*1000,0))",
 'AA43': '=B43',
 'AD43': '=I43+K43+L43',
 'AE43': "='[17]GSNG Imbalance'!Y41",
 'AF43': '=AD43+AE43',
 'C45': '=SUM(C13:C43)',
 'F45': '=SUM(F13:F43)',
 'G45': '=SUM(G13:G43)',
 'H45': '=SUM(H13:H43)',
 'I45': '=SUM(I13:I43)',
 'J45': '=SUM(J13:J43)',
 'K45': '=SUM(K13:K43)',
 'L45': '=SUM(L13:L43)',
 'M45': '=SUM(M13:M43)',
 'N45': '=SUM(N13:N43)',
 'S45': '=AVERAGE(S13:S43)',
 'T45': '=AVERAGE(T13:T43)',
 'U45': '=AVERAGE(U13:U43)',
 'V45': '=AVERAGE(V13:V43)',
 'W45': '=AVERAGE(W13:W43)',
 'Y45': '=AVERAGE(Y13:Y43)',
 'Z45': '=AVERAGE(Z13:Z43)',
 'F47': '=+F45*10',
 'L47': '=Q43'}

def formula_updates():
    """Replace the external book prefix with embedded dependency sheets."""
    return {address: formula.replace("'[17]", "'")
            for address, formula in CLIENT_FORMULAS.items()}
