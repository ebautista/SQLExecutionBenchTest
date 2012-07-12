Attribute VB_Name = "MQueries"
Option Explicit


Public Function GenerateUniqueCode(datData As DAO.Database) As String
    Dim rstTempData As DAO.Recordset
    Dim strUniqueCode As String
    
    Set rstTempData = datData.OpenRecordset("MasterPLDA")
    rstTempData.Index = "CODE"
    
    Do While (rstTempData.NoMatch = False)

        strUniqueCode = TheUniqueCode()
        rstTempData.Seek "=", strUniqueCode
    Loop
    
    rstTempData.Close
    Set rstTempData = Nothing
    
    GenerateUniqueCode = strUniqueCode
    
End Function


Function TheUniqueCode() As String

    Dim cUniqueC As String, i As Byte
    Randomize
    cUniqueC = Trim(CStr(CDbl(999999999999998# * Rnd + 1)))
    
    For i = 1 To Len(cUniqueC)
      If InStr(1, "0123456789", Mid(cUniqueC, i, 1)) > 0 Then
        TheUniqueCode = TheUniqueCode + Mid(cUniqueC, i, 1)
      End If
    Next
    
    TheUniqueCode = Repl("0", 21 - Len(TheUniqueCode)) + TheUniqueCode
End Function


Public Function Repl(ByVal StringToReplicate As String, ByVal HowManyTimes As Integer) As String

    Dim intCtr As Integer
    Dim strReplicatedString As String
    
    For intCtr = 1 To HowManyTimes
        strReplicatedString = strReplicatedString & StringToReplicate
    Next
    
    Repl = strReplicatedString
End Function


Public Sub ExecuteADOInsertsImport(conSadbel As ADODB.Connection, UniqueCode As String)
    
    Dim strCommand As String
    
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "INSERT INTO [PLDA IMPORT HEADER] (A1, [Book Name], A2, A9, AC, A3, A4, A5, A6, A8, AA, AB, D1, D2, D3, D4, DA, DG, C1, C7, C2, C3, D5, D7, D8, D9, C4, C5, C6, C9, CA, CB, B1, B4, B5, B2, B6, [Code], [Header]) VALUES ('IM', '', 'Z', 'NL', 'VOS 1', 'P241675822300110000012', '20120710', 'OEVEL', 'BEHSS216000', '0', '', '', '35', '35', '23', 'BEHSS216000', 'CN', '', '611.96', '11', 'EXW', 'China', '1', 'BE', '1', '1', '28.75', 'EUR', '1', '', '', '', 'A', '', '851AB8902', 'A', '', '" & UniqueCode & "', 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER ZEGELS] (E1, E2, E3, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0000894346433', '', '', '', '', 'ACCOUNTANT DANNEELS = VICTORINOX', 'NEERHOFSTRAAT 24', '', '2180', 'ANTWERPEN', 'ANTWERPEN', 'BE', 'CLAESSWINNEN JOS', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0000894346433', '', '', '', '', 'ACCOUNTANT DANNEELS = VICTORINOX', 'NEERHOFSTRAAT 24', '', '2180', 'ANTWERPEN', 'ANTWERPEN', 'BE', 'CLAESSWINNEN JOS', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('2', 'BE0000416758223', '1', '8300', '', '', 'VOS LOGISTICS BELGIE NV', 'NIJVERHEIDSSTRAAT 8', '', '2260', 'ANTWERPEN', 'OEVEL', 'BE', 'JOS CLAESSWINNEN', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 2)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 3)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('4', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 4)"
    conSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL] (L1, L2, L3, L4, L5, L6, L7, LC, L8, L9, LA, N1, N2, N3, ND, NE, N4, NF, NG, NH, N5, N9, N7, N8, NB, S1, S2, S3, SF, M1, M2, M3, M4, M5, O5, O6, OB, O7, O8, OC, O9, OA, OD, O1, O2, O3, O4, T1, T2, R1, R2, R3, R5, R6, R8, R9, T3, T4, T5, T6, T7, [Code], [Header], [Detail]) VALUES ('0000000000', '', '', '', '', '', '', '', 'Globalisatie levering in België   week 31   2011 BTW', '35', '35', '40', '71', '', '', '', 'H', '4A0', '', '', '100', '', 'QU', '1', 'QU', 'NG', '23', '1', '', '', '', 'D', 'BE-B-2234', 'BE', '28.75', 'EUR', '1', '', '', '', '', '', '', '1', '', '', '', '', '', 'Z', 'CLE', '20101118', '201044', '', '', 'BEHSS216000', '', '', '', '', 's', '" & UniqueCode & "', 1, " & lngIdx & ")"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL CONTAINER] (S4, S5, S6, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('C600', '2234', '20100922', '', '', '', '', '', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL BIJZONDERE] (P1, P2, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('44-552I001-75', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL ZELF] (U1, U2, U3, [Code], [Header], [Detail], [Ordinal]) VALUES ('A00', '24.11', 'v', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL ZELF] (U1, U2, U3, [Code], [Header], [Detail], [Ordinal]) VALUES ('B00', '139.61', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('1', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 3)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('4', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 4)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] (TZ, T8, T9, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteDAOInsertsImport(datSadbel As DAO.Database, UniqueCode As String)
    
    Dim strCommand As String
    
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "INSERT INTO [PLDA IMPORT HEADER] (A1, [Book Name], A2, A9, AC, A3, A4, A5, A6, A8, AA, AB, D1, D2, D3, D4, DA, DG, C1, C7, C2, C3, D5, D7, D8, D9, C4, C5, C6, C9, CA, CB, B1, B4, B5, B2, B6, [Code], [Header]) VALUES ('IM', '', 'Z', 'NL', 'VOS 1', 'P241675822300110000012', '20120710', 'OEVEL', 'BEHSS216000', '0', '', '', '35', '35', '23', 'BEHSS216000', 'CN', '', '611.96', '11', 'EXW', 'China', '1', 'BE', '1', '1', '28.75', 'EUR', '1', '', '', '', 'A', '', '851AB8902', 'A', '', '" & UniqueCode & "', 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER ZEGELS] (E1, E2, E3, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0000894346433', '', '', '', '', 'ACCOUNTANT DANNEELS = VICTORINOX', 'NEERHOFSTRAAT 24', '', '2180', 'ANTWERPEN', 'ANTWERPEN', 'BE', 'CLAESSWINNEN JOS', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0000894346433', '', '', '', '', 'ACCOUNTANT DANNEELS = VICTORINOX', 'NEERHOFSTRAAT 24', '', '2180', 'ANTWERPEN', 'ANTWERPEN', 'BE', 'CLAESSWINNEN JOS', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('2', 'BE0000416758223', '1', '8300', '', '', 'VOS LOGISTICS BELGIE NV', 'NIJVERHEIDSSTRAAT 8', '', '2260', 'ANTWERPEN', 'OEVEL', 'BE', 'JOS CLAESSWINNEN', '014/578942', '014/594566', 'jclaesswinnen@voslogistics.com', '" & UniqueCode & "', 1, 2)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 3)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA IMPORT HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('4', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 4)"
    datSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL] (L1, L2, L3, L4, L5, L6, L7, LC, L8, L9, LA, N1, N2, N3, ND, NE, N4, NF, NG, NH, N5, N9, N7, N8, NB, S1, S2, S3, SF, M1, M2, M3, M4, M5, O5, O6, OB, O7, O8, OC, O9, OA, OD, O1, O2, O3, O4, T1, T2, R1, R2, R3, R5, R6, R8, R9, T3, T4, T5, T6, T7, [Code], [Header], [Detail]) VALUES ('0000000000', '', '', '', '', '', '', '', 'Globalisatie levering in België   week 31   2011 BTW', '35', '35', '40', '71', '', '', '', 'H', '4A0', '', '', '100', '', 'QU', '1', 'QU', 'NG', '23', '1', '', '', '', 'D', 'BE-B-2234', 'BE', '28.75', 'EUR', '1', '', '', '', '', '', '', '1', '', '', '', '', '', 'Z', 'CLE', '20101118', '201044', '', '', 'BEHSS216000', '', '', '', '', 's', '" & UniqueCode & "', 1, " & lngIdx & ")"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL CONTAINER] (S4, S5, S6, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('C600', '2234', '20100922', '', '', '', '', '', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL BIJZONDERE] (P1, P2, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('44-552I001-75', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL ZELF] (U1, U2, U3, [Code], [Header], [Detail], [Ordinal]) VALUES ('A00', '24.11', 'v', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL ZELF] (U1, U2, U3, [Code], [Header], [Detail], [Ordinal]) VALUES ('B00', '139.61', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('1', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 3)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('4', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 4)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] (TZ, T8, T9, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteADOUpdatesImport(conSadbel As ADODB.Connection, UniqueCode As String)
    
    Dim strCommand As String
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "UPDATE [PLDA IMPORT HEADER] SET A1 = 'IM', [Book Name] = '', A2 = 'Z', A9 = 'NL', AC = 'VOS 1', A3 = 'P241675822300110000012', A4 = '20120710', A5 = 'OEVEL', A6 = 'BEHSS216000', A8 = '0', AA = '', AB = '', D1 = '35', D2 = '35', D3 = '23', D4 = 'BEHSS216000', DA = 'CN', DG = '', C1 = '611.96', C7 = '11', C2 = 'EXW', C3 = 'China', D5 = '1', D7 = 'BE', D8 = '1', D9 = '1', C4 = '28.75', C5 = 'EUR', C6 = '1', C9 = '', CA = '', CB = '', B1 = 'A', B4 = '', B5 = '851AB8902', B2 = 'A', B6 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1"
    conSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER ZEGELS] SET E1 = '', E2 = '', E3 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 1"
    conSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '1', X1 = 'BE0000894346433', XF = '', XD = '', XG = '', XH = '', X2 = 'ACCOUNTANT DANNEELS = VICTORINOX', X3 = 'NEERHOFSTRAAT 24', X4 = '', X5 = '2180', X7 = 'ANTWERPEN', X6 = 'ANTWERPEN', X8 = 'BE', X9 = 'CLAESSWINNEN JOS', XA = '014/578942', XB = '014/594566', XC = 'jclaesswinnen@voslogistics.com'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 1"
    conSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '2', X1 = 'BE0000416758223', XF = '1', XD = '8300', XG = '', XH = '', X2 = 'VOS LOGISTICS BELGIE NV', X3 = 'NIJVERHEIDSSTRAAT 8', X4 = '', X5 = '2260', X7 = 'ANTWERPEN', X6 = 'OEVEL', X8 = 'BE', X9 = 'JOS CLAESSWINNEN', XA = '014/578942', XB = '014/594566', XC = 'jclaesswinnen@voslogistics.com'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 2"
    conSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '3', X1 = '', XF = '', XD = '', XG = '', XH = '', X2 = '', X3 = '', X4 = '', X5 = '', X7 = '', X6 = '', X8 = '', X9 = '', XA = '', XB = '', XC = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 3"
    conSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '4', X1 = '', XF = '', XD = '', XG = '', XH = '', X2 = '', X3 = '', X4 = '', X5 = '', X7 = '', X6 = '', X8 = '', X9 = '', XA = '', XB = '', XC = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 4"
    conSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "UPDATE [PLDA IMPORT DETAIL] SET L1 = '0000000000', L2 = '', L3 = '', L4 = '', L5 = '', L6 = '', L7 = '', LC = '', L8 = 'Globalisatie levering in België   week 31   2011 BTW', L9 = '35', LA = '35', N1 = '40', N2 = '71', N3 = '', ND = '', NE = '', N4 = 'H', NF = '4A0', NG = '', NH = '', N5 = '100', N9 = '', N7 = 'QU', N8 = '1', NB = 'QU', S1 = 'NG', S2 = '23', S3 = '1', SF = '', M1 = '', M2 = '', M3 = 'D', M4 = 'BE-B-2234', M5 = 'BE', O5 = '28.75', O6 = 'EUR', OB = '1', O7 = '', O8 = '', OC = '', O9 = '', OA = '', OD = '', O1 = '1', O2 = '', O3 = '', O4 = '', T1 = '', T2 = '', R1 = 'Z', R2 = 'CLE', R3 = '20101118', R5 = '201044', R6 = '', R8 = '', R9 = 'BEHSS216000', T3 = '', T4 = '', T5 = '', T6 = '', T7 = 's'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL CONTAINER] SET S4 = '', S5 = '', S6 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1 "
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL DOCUMENTEN] SET Q1 = 'C600', Q2 = '2234', Q3 = '20100922', Q4 = '', QB = '', QC = '', Q5 = '', Q7 = '', Q8 = '', Q9 = '', QA = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL BIJZONDERE] SET P1 = '44-552I001-75', P2 = '', P5 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL ZELF] SET U1 = 'A00', U2 = '24.11', U3 = 'v'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL ZELF] SET U1 = 'B00', U2 = '139.61', U3 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 2"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '1', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '2', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 2"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '3', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 3"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '4', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 4"
        conSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] SET TZ = '', T8 = '', T9 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        conSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteDAOUpdatesImport(datSadbel As DAO.Database, UniqueCode As String)
    
    Dim strCommand As String
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "UPDATE [PLDA IMPORT HEADER] SET A1 = 'IM', [Book Name] = '', A2 = 'Z', A9 = 'NL', AC = 'VOS 1', A3 = 'P241675822300110000012', A4 = '20120710', A5 = 'OEVEL', A6 = 'BEHSS216000', A8 = '0', AA = '', AB = '', D1 = '35', D2 = '35', D3 = '23', D4 = 'BEHSS216000', DA = 'CN', DG = '', C1 = '611.96', C7 = '11', C2 = 'EXW', C3 = 'China', D5 = '1', D7 = 'BE', D8 = '1', D9 = '1', C4 = '28.75', C5 = 'EUR', C6 = '1', C9 = '', CA = '', CB = '', B1 = 'A', B4 = '', B5 = '851AB8902', B2 = 'A', B6 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1"
    datSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER ZEGELS] SET E1 = '', E2 = '', E3 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 1"
    datSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '1', X1 = 'BE0000894346433', XF = '', XD = '', XG = '', XH = '', X2 = 'ACCOUNTANT DANNEELS = VICTORINOX', X3 = 'NEERHOFSTRAAT 24', X4 = '', X5 = '2180', X7 = 'ANTWERPEN', X6 = 'ANTWERPEN', X8 = 'BE', X9 = 'CLAESSWINNEN JOS', XA = '014/578942', XB = '014/594566', XC = 'jclaesswinnen@voslogistics.com'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 1"
    datSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '2', X1 = 'BE0000416758223', XF = '1', XD = '8300', XG = '', XH = '', X2 = 'VOS LOGISTICS BELGIE NV', X3 = 'NIJVERHEIDSSTRAAT 8', X4 = '', X5 = '2260', X7 = 'ANTWERPEN', X6 = 'OEVEL', X8 = 'BE', X9 = 'JOS CLAESSWINNEN', XA = '014/578942', XB = '014/594566', XC = 'jclaesswinnen@voslogistics.com'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 2"
    datSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '3', X1 = '', XF = '', XD = '', XG = '', XH = '', X2 = '', X3 = '', X4 = '', X5 = '', X7 = '', X6 = '', X8 = '', X9 = '', XA = '', XB = '', XC = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 3"
    datSadbel.Execute strCommand
    
    strCommand = "UPDATE [PLDA IMPORT HEADER HANDELAARS] SET XE = '4', X1 = '', XF = '', XD = '', XG = '', XH = '', X2 = '', X3 = '', X4 = '', X5 = '', X7 = '', X6 = '', X8 = '', X9 = '', XA = '', XB = '', XC = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Ordinal] = 4"
    datSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "UPDATE [PLDA IMPORT DETAIL] SET L1 = '0000000000', L2 = '', L3 = '', L4 = '', L5 = '', L6 = '', L7 = '', LC = '', L8 = 'Globalisatie levering in België   week 31   2011 BTW', L9 = '35', LA = '35', N1 = '40', N2 = '71', N3 = '', ND = '', NE = '', N4 = 'H', NF = '4A0', NG = '', NH = '', N5 = '100', N9 = '', N7 = 'QU', N8 = '1', NB = 'QU', S1 = 'NG', S2 = '23', S3 = '1', SF = '', M1 = '', M2 = '', M3 = 'D', M4 = 'BE-B-2234', M5 = 'BE', O5 = '28.75', O6 = 'EUR', OB = '1', O7 = '', O8 = '', OC = '', O9 = '', OA = '', OD = '', O1 = '1', O2 = '', O3 = '', O4 = '', T1 = '', T2 = '', R1 = 'Z', R2 = 'CLE', R3 = '20101118', R5 = '201044', R6 = '', R8 = '', R9 = 'BEHSS216000', T3 = '', T4 = '', T5 = '', T6 = '', T7 = 's'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL CONTAINER] SET S4 = '', S5 = '', S6 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1 "
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL DOCUMENTEN] SET Q1 = 'C600', Q2 = '2234', Q3 = '20100922', Q4 = '', QB = '', QC = '', Q5 = '', Q7 = '', Q8 = '', Q9 = '', QA = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL BIJZONDERE] SET P1 = '44-552I001-75', P2 = '', P5 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL ZELF] SET U1 = 'A00', U2 = '24.11', U3 = 'v'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL ZELF] SET U1 = 'B00', U2 = '139.61', U3 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 2"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '1', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '2', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 2"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '3', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 3"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL HANDELAARS] SET VE = '4', V1 = '', VG = '', VH = '', V2 = '', V3 = '', V4 = '', V5 = '', V7 = '', V6 = '', V8 = ''WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 4"
        datSadbel.Execute strCommand
        
        strCommand = "UPDATE [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] SET TZ = '', T8 = '', T9 = 'E'WHERE [Code] = '" & UniqueCode & "' AND [Header] = 1 AND [Detail] = " & lngIdx & " AND [Ordinal] = 1"
        datSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteCleanupADOImport(conSadbel As ADODB.Connection, arrUniqueCode() As String)
    
    Dim UniqueCode As Variant
    
    For Each UniqueCode In arrUniqueCode
        conSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER ZEGELS] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL CONTAINER] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL DOCUMENTEN] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL BIJZONDERE] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL ZELF] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] WHERE CODE = '" & UniqueCode & "'"
    Next UniqueCode
End Sub


Public Sub ExecuteCleanupDAOImport(datSadbel As DAO.Database, arrUniqueCode() As String)
    
    Dim UniqueCode As Variant
    
    For Each UniqueCode In arrUniqueCode
        datSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER ZEGELS] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT HEADER HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL CONTAINER] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL DOCUMENTEN] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL BIJZONDERE] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL ZELF] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA IMPORT DETAIL BEREKENINGS EENHEDEN] WHERE CODE = '" & UniqueCode & "'"
    Next UniqueCode
End Sub

Public Sub ExecuteADOInsertsCombined(conSadbel As ADODB.Connection, UniqueCode As String)
    
    Dim strCommand As String
    Dim strMessage As String
    
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "INSERT INTO [PLDA COMBINED HEADER] (A1, [Book Name], A2, AD, A9, AC, A3, A4, A5, A6, A7, A8, AA, AB, AH, AI, AJ, AK, AL, AM, AN, D1, D2, D3, C4, C5, C6, C2, C3, C7, D5, D8, D9, D6, D7, DF, D4, DB, DC, DD, DE, DG, AO, [Code], [Header]) VALUES ('EX', '', 'Z', '', 'NL', 'TRAD 1', 'P244827589700110000003', '20120712', 'UIKHOVEN', 'BEHSS216001', 'BE101000', '0', '', 'METROTILE.VLG.696', '', '', '', 'A3', '', '', '', '26440.000', '26100.000', '19', '40811.8', 'USD', '1.3387', 'CIF', 'TINCAN ISLAND', '11', 'ECMU 460479-4', '1', '3', 'ECMU 460479-4', '', '', 'BETONZ2224001', 'BE', '', '', '', '', '', '" & UniqueCode & "', 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER ZEKERHEID] (E4, E5, E6, E7, E8, E9, EA, EB, EC, ED, EE, EF, [Code], [Header], [Ordinal]) VALUES ('', '', '', '', '', '', '', '', '', '', '', 'E', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER ZEGELS] (E1, E2, E3, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0050448275897', '1', '2128', '', '', '', '', '', '', '', '', '', 'TRADUBEL SA', '087851113', '087866786', 'EYNATTEN@TRADUBEL.BE', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 2)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 3)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('4', '', '', '', '', '', 'METROTILE', 'J.G. STREET', '', '999', '', 'ABUJA', 'NG', '', '', '', '', '" & UniqueCode & "', 1, 4)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('5', 'BE0050460680219', '', '2224', '', '', '', '', '', '', '', '', '', 'SARAH', '012241801', '012241802', '', '" & UniqueCode & "', 1, 5)"
    conSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER TRANSIT OFFICES] (AE, AF, AG, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    conSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL] (L1, L2, L3, L4, L5, L6, LC, L8, L9, LA, N1, N2, N3, ND, NE, N4, NF, NG, NH, N9, N7, NB, NC, S1, S2, S3, SF, M1, M2, O3, O4, M3, M4, M5, O2, O6, OB, R1, R2, R3, R5, R6, R8, R9, T7, [Code], [Header], [Detail]) VALUES ('68079000', '', '', '', '', '', '', 'werken van asfalt of van dergelijke producten b.v. petroleumbitumen, koolteerpek (m.u.v. die op rollen)', '26440.000', '26100.000', '10', '0', '', '', '', 'A', '', '', '', '', 'NG', '', '1', 'PX', '19', 'PX', '', '', '', '', '', '', '', '', '21400', 'EUR', '1', 'Z', 'ZZZ', '20111206', 'VLG.696', '', '', '', 'F', '" & UniqueCode & "', 1, " & lngIdx & ")"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL CONTAINER] (S4, S5, S6, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 2, 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('N380', '12455', '20111206', '', '', '', '1', '', '', '', 'V', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('3028', '2224', '20111206', '', '', '', '1', '', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL BIJZONDERE] (P1, P2, P3, P4, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('30400', '', '', '', 'V', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL BIJZONDERE] (P1, P2, P3, P4, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('ALGEN06', 'BE0460680219', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('1', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 3)"
        conSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL SENSITIVE GOODS] (SB, SC, SD, SE, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        conSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteDAOInsertsCombined(datSadbel As DAO.Database, UniqueCode As String)
    
    Dim strCommand As String
    Dim strMessage As String
    
    Dim lngIdx As Long
    
    '*********************************************************************************************************************************
    'HEADERS
    '*********************************************************************************************************************************
    strCommand = "INSERT INTO [PLDA COMBINED HEADER] (A1, [Book Name], A2, AD, A9, AC, A3, A4, A5, A6, A7, A8, AA, AB, AH, AI, AJ, AK, AL, AM, AN, D1, D2, D3, C4, C5, C6, C2, C3, C7, D5, D8, D9, D6, D7, DF, D4, DB, DC, DD, DE, DG, AO, [Code], [Header]) VALUES ('EX', '', 'Z', '', 'NL', 'TRAD 1', 'P244827589700110000003', '20120712', 'UIKHOVEN', 'BEHSS216001', 'BE101000', '0', '', 'METROTILE.VLG.696', '', '', '', 'A3', '', '', '', '26440.000', '26100.000', '19', '40811.8', 'USD', '1.3387', 'CIF', 'TINCAN ISLAND', '11', 'ECMU 460479-4', '1', '3', 'ECMU 460479-4', '', '', 'BETONZ2224001', 'BE', '', '', '', '', '', '" & UniqueCode & "', 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER ZEKERHEID] (E4, E5, E6, E7, E8, E9, EA, EB, EC, ED, EE, EF, [Code], [Header], [Ordinal]) VALUES ('', '', '', '', '', '', '', '', '', '', '', 'E', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER ZEGELS] (E1, E2, E3, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('1', 'BE0050448275897', '1', '2128', '', '', '', '', '', '', '', '', '', 'TRADUBEL SA', '087851113', '087866786', 'EYNATTEN@TRADUBEL.BE', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 2)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, 3)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('4', '', '', '', '', '', 'METROTILE', 'J.G. STREET', '', '999', '', 'ABUJA', 'NG', '', '', '', '', '" & UniqueCode & "', 1, 4)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER HANDELAARS] (XE, X1, XF, XD, XG, XH, X2, X3, X4, X5, X7, X6, X8, X9, XA, XB, XC, [Code], [Header], [Ordinal]) VALUES ('5', 'BE0050460680219', '', '2224', '', '', '', '', '', '', '', '', '', 'SARAH', '012241801', '012241802', '', '" & UniqueCode & "', 1, 5)"
    datSadbel.Execute strCommand
    
    strCommand = "INSERT INTO [PLDA COMBINED HEADER TRANSIT OFFICES] (AE, AF, AG, [Code], [Header], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 1)"
    datSadbel.Execute strCommand
    '*********************************************************************************************************************************
    
    '*********************************************************************************************************************************
    'DETAILS
    '*********************************************************************************************************************************
    For lngIdx = 0 To FMain.txtDetails.Text
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL] (L1, L2, L3, L4, L5, L6, LC, L8, L9, LA, N1, N2, N3, ND, NE, N4, NF, NG, NH, N9, N7, NB, NC, S1, S2, S3, SF, M1, M2, O3, O4, M3, M4, M5, O2, O6, OB, R1, R2, R3, R5, R6, R8, R9, T7, [Code], [Header], [Detail]) VALUES ('68079000', '', '', '', '', '', '', 'werken van asfalt of van dergelijke producten b.v. petroleumbitumen, koolteerpek (m.u.v. die op rollen)', '26440.000', '26100.000', '10', '0', '', '', '', 'A', '', '', '', '', 'NG', '', '1', 'PX', '19', 'PX', '', '', '', '', '', '', '', '', '21400', 'EUR', '1', 'Z', 'ZZZ', '20111206', 'VLG.696', '', '', '', 'F', '" & UniqueCode & "', 1, " & lngIdx & ")"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL CONTAINER] (S4, S5, S6, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', 'E', '" & UniqueCode & "', 1, 2, 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('N380', '12455', '20111206', '', '', '', '1', '', '', '', 'V', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL DOCUMENTEN] (Q1, Q2, Q3, Q4, QB, QC, Q5, Q7, Q8, Q9, QA, [Code], [Header], [Detail], [Ordinal]) VALUES ('3028', '2224', '20111206', '', '', '', '1', '', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL BIJZONDERE] (P1, P2, P3, P4, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('30400', '', '', '', 'V', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL BIJZONDERE] (P1, P2, P3, P4, P5, [Code], [Header], [Detail], [Ordinal]) VALUES ('ALGEN06', 'BE0460680219', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('1', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('2', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 2)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL HANDELAARS] (VE, V1, VG, VH, V2, V3, V4, V5, V7, V6, V8, [Code], [Header], [Detail], [Ordinal]) VALUES ('3', '', '', '', '', '', '', '', '', '', '', '" & UniqueCode & "', 1, " & lngIdx & ", 3)"
        datSadbel.Execute strCommand
        
        strCommand = "INSERT INTO [PLDA COMBINED DETAIL SENSITIVE GOODS] (SB, SC, SD, SE, [Code], [Header], [Detail], [Ordinal]) VALUES ('', '', '', 'E', '" & UniqueCode & "', 1, " & lngIdx & ", 1)"
        datSadbel.Execute strCommand
    Next lngIdx
    '*********************************************************************************************************************************
End Sub


Public Sub ExecuteCleanupADOCombined(conSadbel As ADODB.Connection, arrUniqueCode() As String)
    
    Dim UniqueCode As Variant
    
    For Each UniqueCode In arrUniqueCode
        conSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER ZEGELS] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER ZEKERHEID] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER TRANSIT OFFICES] WHERE CODE = '" & UniqueCode & "'"
        
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL CONTAINER] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL DOCUMENTEN] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL BIJZONDERE] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL SENSITIVE GOODS] WHERE CODE = '" & UniqueCode & "'"
        conSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
    Next UniqueCode
End Sub


Public Sub ExecuteCleanupDAOCombined(datSadbel As DAO.Database, arrUniqueCode() As String)
    
    Dim UniqueCode As Variant
    
    For Each UniqueCode In arrUniqueCode
        datSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER ZEGELS] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER ZEKERHEID] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED HEADER TRANSIT OFFICES] WHERE CODE = '" & UniqueCode & "'"
        
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL CONTAINER] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL DOCUMENTEN] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL BIJZONDERE] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL SENSITIVE GOODS] WHERE CODE = '" & UniqueCode & "'"
        datSadbel.Execute "DELETE FROM [PLDA COMBINED DETAIL HANDELAARS] WHERE CODE = '" & UniqueCode & "'"
    Next UniqueCode
End Sub
