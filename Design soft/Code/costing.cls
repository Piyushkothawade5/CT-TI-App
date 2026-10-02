VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "costing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''Dim CRGO As String
''Dim CoreSz As String
''Dim rst As Recordset
''
''Public Function FindCoreID(DesCode As String, CShape As String, CorNo As Integer) As String
''
''    Dim rs As New ADODB.Recordset
''    Dim rs1 As New ADODB.Recordset
''    Dim DWL, WWW, DOL, DOW, DHT As String
''    Dim DID, DOD As String
''    If CShape = "R" Then
''    'For Rectangular
''        Set rs = New ADODB.Recordset
''        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''        If rs.EOF = False And rs.BOF = False Then
''
''            DWw = CStr(rs!WW)
''            DWL = CStr(rs!WL)
''            DOL = CStr(rs!OL)
''            DOW = CStr(rs!OW)
''            DHT = CStr(rs!HT)
''
''            Dim y As Double
''            Dim X As Integer
''
''            Size = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
''
'''            MsgBox Size
''            Dim rst As New ADODB.Recordset
''            Set rst = New ADODB.Recordset
''            rst.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'RECTANGULAR' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
''            If rst.EOF = False And rst.BOF = False Then
''                FindCoreID = rst!product_code
''            End If
''
''        End If
''    Else
''    'For Circular
''        Set rs = New ADODB.Recordset
''        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''        If rs.EOF = False And rs.BOF = False Then
''            DID = CStr(rs!oD)
''            DOD = CStr(rs!ID)
''            DHT = CStr(rs!HT)
''            Size = DOD & "X" & DID & "X" & DHT
''
''            Dim rst1 As ADODB.Recordset
''            Set rst1 = New ADODB.Recordset
''
''            rst1.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'ROUND;' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DOD & "X" & DID & "X" & DHT
''            If rst1.EOF = False And rst1.BOF = False Then
''                FindCoreID = rst1!product_code
''            End If
''            Set rst1 = Nothing
''
''        End If
''    End If
''    Set srt = Nothing
''Exit Function
'''    Dim rs As New ADODB.Recordset
'''    Dim rs1 As New ADODB.Recordset
'''    Dim DWL, WWW, DOL, DOW, DHT As String
'''    Dim DID, DOD As String
'''    If CShape = "R" Then
'''    'For Rectangular
'''
'''        Set rs = New ADODB.Recordset
'''        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
'''        If rs.EOF = False And rs.BOF = False Then
''''            If rs!WW > rs!WL Then
'''                DWw = CStr(rs!WW)
'''                DWL = CStr(rs!WL)
'''
'''                DOL = CStr(rs!OL)
'''                DOW = CStr(rs!OW)
'''
'''                DHT = CStr(rs!HT)
''''            Else
''''                DWL = CStr(rs!WW)
''''                DWW = CStr(rs!WL)
''''
''''                DOL = CStr(rs!OW)
''''                DOW = CStr(rs!OL)
''''
''''                DHT = CStr(rs!HT)
''''            End If
'''            '+++++++++++++++++++++++++Pragati++++++++++++++
'''            Dim y As Double
'''            Dim X As Integer
''''            If Len(DWL) = 4 Then
''''                Size = DWL
''''            ElseIf Len(DWL) = 3 Then
''''                Size = DWL & "0"
''''            ElseIf Len(DWL) = 2 Then
''''                Size = "0" & DWL & "0"
''''            ElseIf Len(DWL) = 1 Then
''''                Size = "00" & DWL & "0"
''''            End If
''''            If Len(DWw) = 4 Then
''''                Size = Size & DWw
''''            ElseIf Len(DWw) = 3 Then
''''                Size = Size & DWw & "0"
''''            ElseIf Len(DWw) = 2 Then
''''                Size = Size & "0" & DWw & "0"
''''            ElseIf Len(DWw) = 1 Then
''''                Size = Size & "00" & DWw & "0"
''''            End If
''''            y = (CDbl(DOL) - CDbl(DWL)) / 2
''''            'txtOOd.Text = CDbl(DWW) + y * 2
''''        '    MsgBox CInt(X)
''''            If InStr(1, CStr(y), ".") > 0 Then
''''                    X = y * 10
''''                Else
''''                    X = y
''''            End If
''''            If Len(CStr(X)) = 4 Then
''''                Size = Size & X
''''            ElseIf Len(CStr(X)) = 3 Then
''''                Size = Size & "0" & X
''''            ElseIf Len(CStr(X)) = 2 Then
''''                Size = Size & "0" & X & "0"
''''            ElseIf Len(CStr(X)) = 1 Then
''''                Size = Size & "00" & X & "0"
''''            End If
''''
''''            If Len(DHT) = 4 Then
''''                Size = Size & DHT
''''            ElseIf Len(DHT) = 3 Then
''''                Size = Size & DHT & "0"
''''            ElseIf Len(DHT) = 2 Then
''''                Size = Size & "0" & DHT & "0"
''''            ElseIf Len(DHT) = 1 Then
''''                Size = Size & "00" & DHT & "0"
''''            End If
'''
'''            '++++++++++++++++++++++++++++++++++++++ Pragati end +++++++++++++++
'''
'''            Size = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
'''
''''            MsgBox Size
'''            Dim rst As New ADODB.Recordset
'''            Set rst = New ADODB.Recordset
'''            rst.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'ROUND;' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'''            CoreSz = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
'''            If rst.EOF = False And rst.BOF = False Then
''''                MsgBox "The item you are searching for is : " & "  " & rst!Item_code
'''                FindCoreID = rst!Product_code
''''            Else
'''
'''                '"""""""
'''                'reversing the size and checking again
''''                DWL = CStr(rs!WW)
''''                DWw = CStr(rs!WL)
''''
''''                DOL = CStr(rs!OW)
''''                DOW = CStr(rs!OL)
''''
''''                DHT = CStr(rs!HT)
''''                If Len(DWL) = 4 Then
''''                    Size = DWL
''''                ElseIf Len(DWL) = 3 Then
''''                    Size = DWL & "0"
''''                ElseIf Len(DWL) = 2 Then
''''                    Size = "0" & DWL & "0"
''''                ElseIf Len(DWL) = 1 Then
''''                    Size = "00" & DWL & "0"
''''                End If
''''
''''                If Len(DWw) = 4 Then
''''                    Size = Size & DWw
''''                ElseIf Len(DWw) = 3 Then
''''                    Size = Size & DWw & "0"
''''                ElseIf Len(DWw) = 2 Then
''''                    Size = Size & "0" & DWw & "0"
''''                ElseIf Len(DWw) = 1 Then
''''                    Size = Size & "00" & DWw & "0"
''''                End If
''''
''''
''''                y = (CDbl(DOL) - CDbl(DWL)) / 2
'''                'txtOOd.Text = CDbl(DWW) + y * 2
'''            '    MsgBox CInt(X)
''''                If InStr(1, CStr(y), ".") > 0 Then
''''                        X = y * 10
''''                    Else
''''                        X = y
''''                End If
''''                If Len(CStr(X)) = 4 Then
''''                    Size = Size & X
''''                ElseIf Len(CStr(X)) = 3 Then
''''                    Size = Size & "0" & X
''''                ElseIf Len(CStr(X)) = 2 Then
''''                    Size = Size & "0" & X & "0"
''''                ElseIf Len(CStr(X)) = 1 Then
''''                    Size = Size & "00" & X & "0"
''''                End If
''''
''''                If Len(DHT) = 4 Then
''''                    Size = Size & DHT
''''                ElseIf Len(DHT) = 3 Then
''''                    Size = Size & DHT & "0"
''''                ElseIf Len(DHT) = 2 Then
''''                    Size = Size & "0" & DHT & "0"
''''                ElseIf Len(DHT) = 1 Then
''''                    Size = Size & "00" & DHT & "0"
''''                End If
''''
'''''                MsgBox Size
''''                'Dim rst As New ADODB.Recordset
''''                Set rst = New ADODB.Recordset
''''                rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''''                CoreSz = DWw & "X" & DWL & " " & DOW & "X" & DOL & " " & DHT
''''                If rst.EOF = False And rst.BOF = False Then
'''''                    MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''''                    FindCoreID = rst!item_code
''''                End If
'''
'''                '"""""""
'''            End If
'''
'''        End If
'''        'If found the Code
'''
'''
'''        'Else generate new
'''
'''    Else
'''    'For Circular
'''        Set rs = New ADODB.Recordset
'''        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
'''
'''        If rs.EOF = False And rs.BOF = False Then
''''                If rs!ID > rs!od Then
'''                    DID = CStr(rs!oD)
'''                    DOD = CStr(rs!ID)
'''                    DHT = CStr(rs!HT)
''''                Else
''''                    DID = CStr(rs!ID)
''''                    DOD = CStr(rs!od)
''''                    DHT = CStr(rs!HT)
''''                End If
'''                Size = ""
'''
'''
'''                If Len(DID) = 4 Then
'''                    Size = DID
'''                ElseIf Len(DID) = 3 Then
'''                    Size = DID & "0"
'''                ElseIf Len(DID) = 2 Then
'''                    Size = "0" & DID & "0"
'''                ElseIf Len(DID) = 1 Then
'''                    Size = "00" & DID & "0"
'''                End If
'''
'''                If Len(DOD) = 4 Then
'''                    Size = Size & DOD
'''                ElseIf Len(DOD) = 3 Then
'''                    Size = Size & DOD & "0"
'''                ElseIf Len(DOD) = 2 Then
'''                    Size = Size & "0" & DOD & "0"
'''                ElseIf Len(DOD) = 1 Then
'''                    Size = Size & "00" & DOD & "0"
'''                End If
'''
'''
'''                If Len(DHT) = 3 Then
'''                    Size = Size & DHT & "0"
'''                ElseIf Len(DHT) = 2 Then
'''                    Size = Size & "0" & DHT & "0"
'''                ElseIf Len(DHT) = 1 Then
'''                    Size = Size & "000" & DHT
'''                End If
'''
''''            MsgBox Size
''''            Dim rst As New ADODB.Recordset
'''            Dim rst1 As ADODB.Recordset
'''            Set rst1 = New ADODB.Recordset
'''            rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'''            CoreSz = DID & " X " & DOD & "  " & DHT
'''            If rst1.EOF = False And rst1.BOF = False Then
''''                MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
'''                FindCoreID = rst1!item_code
'''            Else
'''                'Reversing Size and checking again
'''                    DID = CStr(rs!ID)
'''                    DOD = CStr(rs!oD)
'''                    DHT = CStr(rs!HT)
'''                    Size = ""
'''
'''                    If Len(DID) = 4 Then
'''                        Size = DID
'''                    ElseIf Len(DID) = 3 Then
'''                        Size = DID & "0"
'''                    ElseIf Len(DID) = 2 Then
'''                        Size = "0" & DID & "0"
'''                    ElseIf Len(DID) = 1 Then
'''                        Size = "00" & DID & "0"
'''                    End If
'''
'''                    If Len(DOD) = 4 Then
'''                        Size = Size & DOD
'''                    ElseIf Len(DOD) = 3 Then
'''                        Size = Size & DOD & "0"
'''                    ElseIf Len(DOD) = 2 Then
'''                        Size = Size & "0" & DOD & "0"
'''                    ElseIf Len(DOD) = 1 Then
'''                        Size = Size & "00" & DOD & "0"
'''                    End If
'''
'''
'''                    If Len(DHT) = 3 Then
'''                        Size = Size & DHT & "0"
'''                    ElseIf Len(DHT) = 2 Then
'''                        Size = Size & "0" & DHT & "0"
'''                    ElseIf Len(DHT) = 1 Then
'''                        Size = Size & "000" & DHT
'''                    End If
'''
''''                MsgBox Size
'''    '            Dim rst As New ADODB.Recordset
'''                'Dim rst1 As ADODB.Recordset
'''                Set rst1 = New ADODB.Recordset
'''                rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'''                CoreSz = DID & " X " & DOD & "  " & DHT
'''                If rst1.EOF = False And rst1.BOF = False Then
''''                    MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
'''                    FindCoreID = rst1!item_code
'''                End If
'''                'End Reversing Size and checking again
'''            End If
'''            Set rst1 = Nothing
'''
'''
'''        End If
'''    End If
''
''End Function
''
''Public Function Coating(DES As String)
''Dim EqType As String
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++
'''Deleteing coasting
''Set rst = New Recordset
''rst.Open "delete from costing", con, adOpenDynamic, adLockOptimistic
'' Set rst = Nothing
''
''Dim rst1 As Recordset
''
''Dim priCross As String
''Dim PTurns As Integer
''Dim CopTer1 As String
''Dim CopTer2 As String
''Dim Bar1 As String
''Dim Bar2 As String
''Dim Brackett As String
''Dim Plate As String
''Dim Drg As String
''Dim drgtype As String
''Dim Equip As String
''Dim CoreShape As String
''Dim CoreMat As String
''Dim CoreID As String
''Dim PriCond As String
''Dim PriCondType As String
''Dim PriConductor As String
''
''Dim Bartype As String
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Fing Pric Cross Sction
''Set rst = New Recordset
''rst.Open "select coreshape, pri_cond_style, equip, pri_cross_sec, pri_turns , drg_no , drg_type, pri_cond from designspec where des_code = '" & DES & "'", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''    priCross = rst!pri_cross_sec
''    PTurns = rst!pri_turns
''    Drg = rst!Drg_no & vbNullString
''    drgtype = rst!DRG_TYPE & vbNullString
''    Equip = rst!Equip & vbNullString
''    CoreShape = rst!CoreShape & vbNullString
''    PriConductor = rst!pri_cond & vbNullString
''    'Only primary condutor involved
''    If IsNull(rst!pri_cond_style) = False And Len(Trim(rst!pri_cond_style)) > 0 Then
''        PriCond = Mid(rst!pri_cond_style, 1, 1)
''        PriCondType = Mid(rst!pri_cond_style, InStr(1, rst!pri_cond_style, "-") + 1)
''    End If
''End If
''Set rst = Nothing
''
''If InStr(1, Equip, "Wound") > 0 Then
''    EqType = "2"
''ElseIf InStr(1, Equip, "Window") > 0 Then
''    EqType = "1"
''ElseIf InStr(1, Equip, "Bar") > 0 Then
''    EqType = "1"
''End If
''
''
''Dim Single_Or_Double As String
''If InStr(1, PriCondType, "Double") > 0 Then
''    Single_Or_Double = "Double"
''ElseIf InStr(1, PriCondType, "Single") > 0 Then
''    Single_Or_Double = "Single"
''ElseIf InStr(1, PriCondType, "Series") > 0 Then
''    Single_Or_Double = "Series"
''End If
''
''If Single_Or_Double = "Double" Then
''    If InStr(1, PriCondType, "Equal") > 0 Then
''        Single_Or_Double = "Double-Equal"
''    ElseIf InStr(1, PriCondType, "Different") > 0 Then
''        Single_Or_Double = "Double-Different"
''    Else
''        Single_Or_Double = "Double"
''    End If
''End If
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from primarycond11  -by providing c/s and turns
'''++++++++++++++++++++++++++++++
'''+changing ri primarycond from primartcond11
''Set rst = New Recordset
''str1 = "select * from primarycond11 where t_csp = '" & priCross & "' and t_np = '" & PTurns & "' and style = '" & Single_Or_Double & "'"
''If Trim(priCross) = "" Then
''    PriCond = 0
''    rst.Open "select *  from primarycond11 where [t_np] = " & PTurns & " and [type] = " & PriCond & " and [style] = '" & PriCondType & "' and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and [style] = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''Else
'''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and type = '" & PriCond & "%' and style = '" & PriCondType & "%' and drg = '" & Drg & "' and mod = '" & DrgType & "'", con, adOpenStatic, adLockReadOnly
''    strr = "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and mod = '" & drgtype & "' and style = '" & Single_Or_Double & "'"
''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and [mod] = '" & drgtype & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''End If
''If rst.EOF = False And rst.BOF = False Then
''    CopTer1 = rst!t_terp1 & vbNullString
''    CopTer2 = rst!t_terp2 & vbNullString
''    Bar1 = rst!cond_code1 & vbNullString
''    If Len(rst!cond_code2) > 0 And IsNull(rst!cond_code2) = False Then
''        Bar2 = rst!cond_code2
''    End If
''Else
''    Set rst = New Recordset
''    rst.Open "select *  from primarycond11 where t_np = " & PTurns & " and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and t_cond_siz = '" & PriConductor & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''    If rst.EOF = False And rst.BOF = False Then
''        CopTer1 = rst!t_terp1 & vbNullString
''        CopTer2 = rst!t_terp2 & vbNullString
''        Bar1 = rst!cond_code1 & vbNullString
''        If Len(rst!cond_code2) > 0 Or IsNull(rst!cond_code2) = False Then
''            Bar2 = rst!cond_code2
''        End If
''    End If
''End If
''Set rst = Nothing
''
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from MouldVol -by providing drg, type
''Set rst = New Recordset
''If Len(Trim(drgtype)) > 0 Then
''    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "' and t_drgtype = '" & drgtype & "'", con, adOpenStatic, adLockReadOnly
''Else
''    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
''End If
''If rst.EOF = False And rst.BOF = False Then
''    If InStr(1, UCase(Equip), "WOUND") > 0 Then
''        Brackett = rst!t_bktcode & vbNullString  'for wound
''        Plate = rst!t_pltcode & vbNullString     'for wound
''    Else
''        WinBar = rst!t_mtginst & vbNullString   'for window and  Bar
''        Plate = rst!t_pltcode & vbNullString
''        If InStr(1, UCase(Equip), "BAR") > 0 Then
''            Bartype = rst!t_buscode & vbNullString
''        End If
''    End If
''End If
''Set rst = Nothing
''
'''ADD INF TO REPORT
''
''Dim rss As Recordset
''
''''Core material
'''Set rst = New Recordset
'''rst.Open "select distinct core_mat from designspec where des_code = '" & DES & "'", con, adOpenStatic
'''If rst.EOF = False And rst.BOF = False Then
'''    CoreMat = rst!core_mat
'''End If
'''Set rst = Nothing
''
''Dim epoxy As Double
''Dim rstCRGO As Recordset
''
''Set rst = New Recordset
''rst.Open "select distinct core_no, tot_wt_core_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
'''    If CoreMat = "MU metal" Then
'''        '1***************
'''        If CoreShape = "Rectangluar" Then
'''            we = prcCalMUCoreShape("R", rst!tot_wt_core_wt)
'''        Else
'''            we = prcCalMUCoreShape("C", rst!tot_wt_core_wt)
'''        End If
'''    Else
''        '1***************
''
''        Do Until rst.EOF
''            Set rstCRGO = New Recordset
''            rstCRGO.Open "select core_mat, core_wt from designcore where dec_code = '" & DES & "' and coreno = " & rst!core_no & "", con, adOpenStatic, adLockReadOnly
''            If rstCRGO.EOF = False And rstCRGO.BOF = False Then CRGO = rstCRGO.Fields(0) & vbNullString
''
''            If CoreShape = "Rectangluar" Then
''                CoreID = FindCoreID(DES, "R", rst!core_no)
''                we = prcCalCoreShape("R", rstCRGO!core_wt, CoreID)
''            Else
''                CoreID = FindCoreID(DES, "C", rst!core_no)
''                we = prcCalCoreShape("C", rstCRGO!core_wt, CoreID)
''            End If
''            Set rstCRGO = Nothing
''            rst.MoveNext
''        Loop
'''    End If
''End If
''Set rst = Nothing
''
'''2************************
'''Super Enabled Copper
''Set rst = New Recordset
''rst.Open "select max(coreno) from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''    X = rst.Fields(0)
''End If
''Set rst = Nothing
''
''    'Finding swg of each core fro design core
''For i = 1 To X
''    Set rst = New Recordset
''    rst.Open "select swg, sec_cu_wt from designcore where dec_code = '" & DES & "' and coreno = " & i & " ", con, adOpenStatic
''    If rst.EOF = False And rst.BOF = False Then
''        y = RateCopper(rst!SWG, rst!sec_cu_wt)
''    End If
''    Set rst = Nothing
''Next
''
'''3***************
'''Copper bunch 1 or 2
''Set rst = New Recordset
''rst.Open "select distinct coreno from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''    vall = rst!coreno
''End If
''Set rst = Nothing
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''    '4**************
''    'Terminal
''    we = RateCal(CopTer1, 1)
''    we = RateCal(CopTer2, 1)
''
''
''
''    '' +++++++++++++
''    ''Copper strip tot_wt_pricu_wt
''    Dim rst2 As Recordset
''    Set rst2 = New Recordset
''    rst2.Open "select tot_wt_pricu_wt  from designspec where des_code = '" & DES & "'", con, adOpenStatic
''    If rst2.EOF = False And rst2.BOF = False Then
''        ''if only one group present all wt. taken by 1
''        If Bar2 = "" Then
''            we = RateCal(Bar1, CDbl(rst2!tot_wt_pricu_wt))
''        ''else wt. distributed
''        Else
''            we = RateCal(Bar1, (CDbl(rst2!tot_wt_pricu_wt) * 0.6))
''            we = RateCal(Bar2, (CDbl(rst2!tot_wt_pricu_wt) * 0.4))
''        End If
''    End If
''    Set rst2 = Nothing
''
''End If
''
''If InStr(1, UCase(Equip), "BAR") > 0 Then
''    we = RateCal(Bartype, 1)
''End If
''
'''++++++++++++++++++++
'''5 Resin
'''Dim epoxy As Double
''Set rst = New Recordset
''rst.Open "select distinct tot_wt_epoxy_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''     we = prcCalResin(CDbl(rst.Fields(0)))
''End If
''Set rst = Nothing
''
''
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''
''    '6**************
''    'Bracket
''    we = RateCal(Brackett, 1)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''Else
''
''    '6**************
''    'for insert
''    we = RateCal(CStr(WinBar), 4)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''End If
''
'''+++++++++++++++++++++++++++++++++
'''Rest of the installation
''Set rst1 = New Recordset
''rst1.Open "costing", con, adOpenDynamic, adLockOptimistic
''rst1.AddNew
''rst1!Item_name = "-"
''rst1.Update
''
''Dim rstItem As Recordset
''
''Set rst = New Recordset
''rst.Open "select PRODUCT_CODE FROM MST_PART WHERE DRG_no = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''    P_CODE = rst.Fields(0)
''End If
''Set rst = Nothing
''
''
''
''
''Set rst = New Recordset
''rst.Open "select * from new_costmaster where PRODUCT_CODE = '" & P_CODE & "'", con, adOpenStatic, adLockReadOnly
'''rst.Open "select * from costingcomm", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''    Do Until rst.EOF
''        rst1.AddNew
''        rst1!code = rst!item_code
''        Set rstItem = New Recordset
''        rstItem.Open "select * from MST_PART where PRODUCT_code = '" & rst!item_code & "'", con, adOpenStatic, adLockReadOnly
''            rst1!Item_name = rstItem!PRODUCT_DESCRIPTION & vbNullString
''            rst1!units = rstItem!uomC & vbNullString
''            If IsNull(rstItem!latest_rate) = True Then
''                rst1!Rate = 0
''            Else
''                rst1!Rate = Format(rstItem!latest_rate, "0.00")
''            End If
''            rst1!qty = Format(Trim(rst!qty), "0.00")
''            rst1!amount = Format(CDbl(rst!qty) * CDbl(rst1!Rate), "0.00")
''        Set rstItem = Nothing
''        rst1.Update
''        rst.MoveNext
''    Loop
''    'Sundary Material
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Sundry Material"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Mould Charges
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Mould Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Packing Charges
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Packing Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''
''End If
''Set rst = Nothing
''
''Set rst1 = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''
''Private Function RateCal(item_code As String, CQty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
''Set rst = New Recordset
'''rst.Open "select * from costmaster where part = '" & Item_code & "'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from mst_part where product_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly
''
''If rst.EOF = False And rst.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!code = item_code
'''      If InStr(rst!item_desc1 & vbNullString, "Insert") > 0 Then
'''        If InStr(rst!item_desc1 & vbNullString, "M.S.") > 0 Then
'''        'M.S.
'''            If IsNull(rst!Type) = True Then
''                rs1!Item_name = rst!product_code & " " & rst!Drg_no
'''            Else
'''                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!Mod & ")" & " Type:" & rst!Type
'''            End If
'''        Else
'''        'Brass
'''            If IsNull(rst!Type) = True Then
'''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")"
'''            Else
'''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
'''            End If
'''        End If
'''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "PLATE") > 0 Then
'''      'Mounding Plate
'''        If IsNull(rst!Type) = True Or Len(rst!Type) = 0 Then
'''            If Len(Trim(rst!Mod)) = 0 Then
'''                rs1!Item_name = "Mounting Plate " & rst!Drg
'''            Else
'''                rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")"
'''            End If
'''        Else
'''            rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
'''        End If
'''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BRACKET") > 0 Then
'''      'Mounting Bracket
'''        If IsNull(rst!Type) = True And Len(Trim(rst!Type)) = 0 Then
'''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")"
'''        Else
'''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
'''        End If
'''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "TERMINAL") > 0 Then
'''      'Mounting Bracket
'''        If IsNull(rst!Type) = True Then
'''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")"
'''        Else
'''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
'''        End If
'''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BUNCHED") > 0 Then
'''      'Mounting Bracket
'''        rs1!Item_name = rst!item_desc2 & " group copper bunch"
'''      Else
'''      'Else
'''        rs1!Item_name = rst!item_desc1
'''        If Len(Trim(rst!item_desc4)) > 0 Or IsNull(rst!item_desc4) = False Then
'''            rs1!Item_name = rs1!Item_name & " " & rst!item_desc4
'''        End If
'''
'''      End If
''
''      rs1!units = rst!uomC & vbNullString
''      rs1!Rate = Val(rst!latest_rate & vbNullString)
''      rs1!qty = CQty
''      rs1!amount = Val(rst!latest_rate & vbNullString) * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''
''End Function
''
''
''Private Function prcCalResin(qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For resin Only
''Set rst = New Recordset
'''rst.Open "select * from costmaster where part = 'INS-00A'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from mst_part where product_code = '11010629-00'", con, adOpenStatic, adLockReadOnly
''
''If rst.EOF = False And rst.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = "11010629-00"
''     rs1!Item_name = "Hot Setting Epoxy Resin Mix"
''     rs1!units = "Kgs"
''     rs1!Rate = rst!latest_rate
''     rs1!qty = CInt(qty)
''     rs1!amount = rst!latest_rate * CInt(qty)
''     rs1.Update
''     Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function prcCalMUCoreShape(TyprC As String, qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
''If TyprC = "R" Then
''    rrr = "CRGO Rectangular Maganetic Core"
''    rcode = "REC-00M"
''    RATEE = 2500
''Else
''    rrr = "CRGO Circular Maganetic Core"
''    rcode = "ROC-00M"
''    RATEE = 2500
''End If
''
''Dim rs1 As Recordset
''
''Set rst = New Recordset
''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function GetCorePrice(CoreShape As String, Qtykg As Double) As Currency
''
''    Dim rs As New ADODB.Recordset
''    Dim SQL1 As String
''    Set rs = New ADODB.Recordset
''    If CoreShape = "R" Then
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRE00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRE00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRE00C"
''        End If
''    Else
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRO00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRO00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRO00C"
''        End If
''    End If
''
''
''    rs.Open "Select * from Core_rate where coretype like '" & CoreShape & "%' and grade = '" & Mid(CRGO, 6, 6) & "'", con, adOpenKeyset, adLockReadOnly
'''    rs.Open "select * from itemmast where item_code = '" & SQL1 & "'", con, adOpenStatic, adLockReadOnly
''    If rs.EOF = False And rs.BOF = False Then
''        GetCorePrice = rs!landedprice
''    End If
''    Set rs = Nothing
''End Function
''
''Private Function prcCalCoreShape(TyprC As String, qty As Double, CoreCode As String)
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
'' If TyprC = "R" Then
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("R", qty)
''Else
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("C", qty)
''End If
''
''Dim rs1 As Recordset
''
''Dim rst1 As ADODB.Recordset
''Set rst1 = New Recordset
'''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''rst1.Open "select * from MST_PART where Product_code = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''
''If rst1.EOF = False And rst1.BOF = False Then
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''Else
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = "-"
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''
''Set rst1 = Nothing
''
''End Function
''
''Private Function RateCopper(item_code As String, CQty As Double)
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
'''item_code = "SWG " & item_code
''Set rst = New Recordset
'''rst.Open "select * from costmaster where title like '" & Item_code & "%'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from MST_PART where material_type = 'SUPER ENAMELED ROUND COPPER WINDING WIRE;' and other_det_type = 'MEDIUM COVERING ;' and Dimension like '" & item_code & "%'", con, adOpenStatic, adLockReadOnly
''
'''str1 = "select * from costmaster where title like '" & Item_code & "%'"
''If rst.EOF = False And rst.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!code = rst!product_code & vbNullString 'rst!part
''      rs1!Item_name = item_code & " Super Enamelled copper wire"
''      rs1!units = rst!uomC & vbNullString
''      rs1!Rate = rst!latest_rate
''      rs1!qty = CQty
''      rs1!amount = rst!latest_rate * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''End Function
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
''__________________________________________________________
''__________________________________________________________
''//// changes
''____________________________________________________________
''Dim CRGO As String
''Dim CoreSz As String
''Dim rst As Recordset
''
''Public Function FindCoreID(DesCode As String, CShape As String, CorNo As Integer) As String
''    Dim rs As New ADODB.Recordset
''    Dim rs1 As New ADODB.Recordset
''    Dim DWL, WWW, DOL, DOW, DHT As String
''    Dim DID, DOD As String
''    If CShape = "R" Then
''    'For Rectangular
''
''        Set rs = New ADODB.Recordset
''        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''        If rs.EOF = False And rs.BOF = False Then
'''            If rs!WW > rs!WL Then
''                DWw = CStr(rs!WW)
''                DWL = CStr(rs!WL)
''
''                DOL = CStr(rs!OL)
''                DOW = CStr(rs!OW)
''
''                DHT = CStr(rs!HT)
'''            Else
'''                DWL = CStr(rs!WW)
'''                DWW = CStr(rs!WL)
'''
'''                DOL = CStr(rs!OW)
'''                DOW = CStr(rs!OL)
'''
'''                DHT = CStr(rs!HT)
'''            End If
''            Dim y As Double
''            Dim X As Integer
''            If Len(DWL) = 4 Then
''                Size = DWL
''            ElseIf Len(DWL) = 3 Then
''                Size = DWL & "0"
''            ElseIf Len(DWL) = 2 Then
''                Size = "0" & DWL & "0"
''            ElseIf Len(DWL) = 1 Then
''                Size = "00" & DWL & "0"
''            End If
''
''            If Len(DWw) = 4 Then
''                Size = Size & DWw
''            ElseIf Len(DWw) = 3 Then
''                Size = Size & DWw & "0"
''            ElseIf Len(DWw) = 2 Then
''                Size = Size & "0" & DWw & "0"
''            ElseIf Len(DWw) = 1 Then
''                Size = Size & "00" & DWw & "0"
''            End If
''
''
''            y = (CDbl(DOL) - CDbl(DWL)) / 2
''            'txtOOd.Text = CDbl(DWW) + y * 2
''        '    MsgBox CInt(X)
''            If InStr(1, CStr(y), ".") > 0 Then
''                    X = y * 10
''                Else
''                    X = y
''            End If
''            If Len(CStr(X)) = 4 Then
''                Size = Size & X
''            ElseIf Len(CStr(X)) = 3 Then
''                Size = Size & "0" & X
''            ElseIf Len(CStr(X)) = 2 Then
''                Size = Size & "0" & X & "0"
''            ElseIf Len(CStr(X)) = 1 Then
''                Size = Size & "00" & X & "0"
''            End If
''
''            If Len(DHT) = 4 Then
''                Size = Size & DHT
''            ElseIf Len(DHT) = 3 Then
''                Size = Size & DHT & "0"
''            ElseIf Len(DHT) = 2 Then
''                Size = Size & "0" & DHT & "0"
''            ElseIf Len(DHT) = 1 Then
''                Size = Size & "00" & DHT & "0"
''            End If
''
'''            MsgBox Size
''            Dim rst As New ADODB.Recordset
''            Set rst = New ADODB.Recordset
''            rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DWL & "X" & DWw & " " & DOL & "X" & DOW & " " & DHT
''            If rst.EOF = False And rst.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''                FindCoreID = rst!item_code
''            Else
''                '"""""""
''                'reversing the size and checking again
''                DWL = CStr(rs!WW)
''                DWw = CStr(rs!WL)
''
''                DOL = CStr(rs!OW)
''                DOW = CStr(rs!OL)
''
''                DHT = CStr(rs!HT)
''                If Len(DWL) = 4 Then
''                    Size = DWL
''                ElseIf Len(DWL) = 3 Then
''                    Size = DWL & "0"
''                ElseIf Len(DWL) = 2 Then
''                    Size = "0" & DWL & "0"
''                ElseIf Len(DWL) = 1 Then
''                    Size = "00" & DWL & "0"
''                End If
''
''                If Len(DWw) = 4 Then
''                    Size = Size & DWw
''                ElseIf Len(DWw) = 3 Then
''                    Size = Size & DWw & "0"
''                ElseIf Len(DWw) = 2 Then
''                    Size = Size & "0" & DWw & "0"
''                ElseIf Len(DWw) = 1 Then
''                    Size = Size & "00" & DWw & "0"
''                End If
''
''
''                y = (CDbl(DOL) - CDbl(DWL)) / 2
''                'txtOOd.Text = CDbl(DWW) + y * 2
''            '    MsgBox CInt(X)
''                If InStr(1, CStr(y), ".") > 0 Then
''                        X = y * 10
''                    Else
''                        X = y
''                End If
''                If Len(CStr(X)) = 4 Then
''                    Size = Size & X
''                ElseIf Len(CStr(X)) = 3 Then
''                    Size = Size & "0" & X
''                ElseIf Len(CStr(X)) = 2 Then
''                    Size = Size & "0" & X & "0"
''                ElseIf Len(CStr(X)) = 1 Then
''                    Size = Size & "00" & X & "0"
''                End If
''
''                If Len(DHT) = 4 Then
''                    Size = Size & DHT
''                ElseIf Len(DHT) = 3 Then
''                    Size = Size & DHT & "0"
''                ElseIf Len(DHT) = 2 Then
''                    Size = Size & "0" & DHT & "0"
''                ElseIf Len(DHT) = 1 Then
''                    Size = Size & "00" & DHT & "0"
''                End If
''
'''                MsgBox Size
''                'Dim rst As New ADODB.Recordset
''                Set rst = New ADODB.Recordset
''                rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''                CoreSz = DWw & "X" & DWL & " " & DOW & "X" & DOL & " " & DHT
''                If rst.EOF = False And rst.BOF = False Then
'''                    MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''                    FindCoreID = rst!item_code
''                End If
''
''                '"""""""
''            End If
''
''        End If
''        'If found the Code
''
''
''        'Else generate new
''
''    Else
''    'For Circular
''        Set rs = New ADODB.Recordset
''        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''
''        If rs.EOF = False And rs.BOF = False Then
'''                If rs!ID > rs!od Then
''                    DID = CStr(rs!oD)
''                    DOD = CStr(rs!ID)
''                    DHT = CStr(rs!HT)
'''                Else
'''                    DID = CStr(rs!ID)
'''                    DOD = CStr(rs!od)
'''                    DHT = CStr(rs!HT)
'''                End If
''                Size = ""
''
''                If Len(DID) = 4 Then
''                    Size = DID
''                ElseIf Len(DID) = 3 Then
''                    Size = DID & "0"
''                ElseIf Len(DID) = 2 Then
''                    Size = "0" & DID & "0"
''                ElseIf Len(DID) = 1 Then
''                    Size = "00" & DID & "0"
''                End If
''
''                If Len(DOD) = 4 Then
''                    Size = Size & DOD
''                ElseIf Len(DOD) = 3 Then
''                    Size = Size & DOD & "0"
''                ElseIf Len(DOD) = 2 Then
''                    Size = Size & "0" & DOD & "0"
''                ElseIf Len(DOD) = 1 Then
''                    Size = Size & "00" & DOD & "0"
''                End If
''
''
''                If Len(DHT) = 3 Then
''                    Size = Size & DHT & "0"
''                ElseIf Len(DHT) = 2 Then
''                    Size = Size & "0" & DHT & "0"
''                ElseIf Len(DHT) = 1 Then
''                    Size = Size & "000" & DHT
''                End If
''
'''            MsgBox Size
'''            Dim rst As New ADODB.Recordset
''            Dim rst1 As ADODB.Recordset
''            Set rst1 = New ADODB.Recordset
''            rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DID & " X " & DOD & "  " & DHT
''            If rst1.EOF = False And rst1.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                FindCoreID = rst1!item_code
''            Else
''                'Reversing Size and checking again
''                    DID = CStr(rs!ID)
''                    DOD = CStr(rs!oD)
''                    DHT = CStr(rs!HT)
''                    Size = ""
''
''                    If Len(DID) = 4 Then
''                        Size = DID
''                    ElseIf Len(DID) = 3 Then
''                        Size = DID & "0"
''                    ElseIf Len(DID) = 2 Then
''                        Size = "0" & DID & "0"
''                    ElseIf Len(DID) = 1 Then
''                        Size = "00" & DID & "0"
''                    End If
''
''                    If Len(DOD) = 4 Then
''                        Size = Size & DOD
''                    ElseIf Len(DOD) = 3 Then
''                        Size = Size & DOD & "0"
''                    ElseIf Len(DOD) = 2 Then
''                        Size = Size & "0" & DOD & "0"
''                    ElseIf Len(DOD) = 1 Then
''                        Size = Size & "00" & DOD & "0"
''                    End If
''
''
''                    If Len(DHT) = 3 Then
''                        Size = Size & DHT & "0"
''                    ElseIf Len(DHT) = 2 Then
''                        Size = Size & "0" & DHT & "0"
''                    ElseIf Len(DHT) = 1 Then
''                        Size = Size & "000" & DHT
''                    End If
''
'''                MsgBox Size
''    '            Dim rst As New ADODB.Recordset
''                'Dim rst1 As ADODB.Recordset
''                Set rst1 = New ADODB.Recordset
''                rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''                CoreSz = DID & " X " & DOD & "  " & DHT
''                If rst1.EOF = False And rst1.BOF = False Then
'''                    MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                    FindCoreID = rst1!item_code
''                End If
''                'End Reversing Size and checking again
''            End If
''            Set rst1 = Nothing
''
''
''        End If
''    End If
''
''End Function
''
''Public Function Coating(DES As String)
''Dim EqType As String
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++
'''Deleteing coasting
''Set rst = New Recordset
''rst.Open "delete from costing", con, adOpenDynamic, adLockOptimistic
'' Set rst = Nothing
''
''Dim rst1 As Recordset
''
''Dim priCross As String
''Dim PTurns As Integer
''Dim CopTer1 As String
''Dim CopTer2 As String
''Dim Bar1 As String
''Dim Bar2 As String
''Dim Brackett As String
''Dim Plate As String
''Dim Drg As String
''Dim drgtype As String
''Dim Equip As String
''Dim CoreShape As String
''Dim CoreMat As String
''Dim CoreID As String
''Dim PriCond As String
''Dim PriCondType As String
''Dim PriConductor As String
''
''Dim Bartype As String
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Fing Pric Cross Sction
''Set rst = New Recordset
''rst.Open "select coreshape, pri_cond_style, equip, pri_cross_sec, pri_turns , drg_no , drg_type, pri_cond from designspec where des_code = '" & DES & "'", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''    priCross = rst!pri_cross_sec
''    PTurns = rst!pri_turns
''    Drg = rst!drg_no & vbNullString
''    drgtype = rst!DRG_TYPE & vbNullString
''    Equip = rst!Equip & vbNullString
''    CoreShape = rst!CoreShape & vbNullString
''    PriConductor = rst!pri_cond & vbNullString
''    If IsNull(rst!pri_cond_style) = False And Len(Trim(rst!pri_cond_style)) > 0 Then
''        PriCond = Mid(rst!pri_cond_style, 1, 1)
''        PriCondType = Mid(rst!pri_cond_style, InStr(1, rst!pri_cond_style, "-") + 1)
''    End If
''End If
''Set rst = Nothing
''
''If InStr(1, Equip, "Wound") > 0 Then
''    EqType = "2"
''ElseIf InStr(1, Equip, "Window") > 0 Then
''    EqType = "1"
''ElseIf InStr(1, Equip, "Bar") > 0 Then
''    EqType = "1"
''End If
''
''
''Dim Single_Or_Double As String
''If InStr(1, PriCondType, "Double") > 0 Then
''    Single_Or_Double = "Double"
''ElseIf InStr(1, PriCondType, "Single") > 0 Then
''    Single_Or_Double = "Single"
''ElseIf InStr(1, PriCondType, "Series") > 0 Then
''    Single_Or_Double = "Series"
''End If
''
''If Single_Or_Double = "Double" Then
''    If InStr(1, PriCondType, "Equal") > 0 Then
''        Single_Or_Double = "Double-Equal"
''    ElseIf InStr(1, PriCondType, "Different") > 0 Then
''        Single_Or_Double = "Double-Different"
''    Else
''        Single_Or_Double = "Double"
''    End If
''End If
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from primarycond11  -by providing c/s and turns
'''++++++++++++++++++++++++++++++
'''+changing ri primarycond from primartcond11
''Set rst = New Recordset
''str1 = "select * from primarycond11 where t_csp = '" & priCross & "' and t_np = '" & PTurns & "' and style = '" & Single_Or_Double & "'"
''If Trim(priCross) = "" Then
''    PriCond = 0
''    rst.Open "select *  from primarycond11 where [t_np] = " & PTurns & " and [type] = " & PriCond & " and [style] = '" & PriCondType & "' and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and [style] = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''Else
'''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and type = '" & PriCond & "%' and style = '" & PriCondType & "%' and drg = '" & Drg & "' and mod = '" & DrgType & "'", con, adOpenStatic, adLockReadOnly
''    strr = "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and mod = '" & drgtype & "' and style = '" & Single_Or_Double & "'"
''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and [mod] = '" & drgtype & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''End If
''If rst.EOF = False And rst.BOF = False Then
''    CopTer1 = rst!t_terp1 & vbNullString
''    CopTer2 = rst!t_terp2 & vbNullString
''    Bar1 = rst!cond_code1 & vbNullString
''    If Len(rst!cond_code2) > 0 And IsNull(rst!cond_code2) = False Then
''        Bar2 = rst!cond_code2
''    End If
''Else
''    Set rst = New Recordset
''    rst.Open "select *  from primarycond11 where t_np = " & PTurns & " and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and t_cond_siz = '" & PriConductor & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''    If rst.EOF = False And rst.BOF = False Then
''        CopTer1 = rst!t_terp1 & vbNullString
''        CopTer2 = rst!t_terp2 & vbNullString
''        Bar1 = rst!cond_code1 & vbNullString
''        If Len(rst!cond_code2) > 0 Or IsNull(rst!cond_code2) = False Then
''            Bar2 = rst!cond_code2
''        End If
''    End If
''End If
''Set rst = Nothing
''
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from MouldVol -by providing drg, type
''Set rst = New Recordset
''If Len(Trim(drgtype)) > 0 Then
''    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "' and t_drgtype = '" & drgtype & "'", con, adOpenStatic, adLockReadOnly
''Else
''    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
''End If
''If rst.EOF = False And rst.BOF = False Then
''    If InStr(1, UCase(Equip), "WOUND") > 0 Then
''        Brackett = rst!t_bktcode & vbNullString  'for wound
''        Plate = rst!t_pltcode & vbNullString     'for wound
''    Else
''        WinBar = rst!t_mtginst & vbNullString   'for window and  Bar
''        Plate = rst!t_pltcode & vbNullString
''        If InStr(1, UCase(Equip), "BAR") > 0 Then
''            Bartype = rst!t_buscode & vbNullString
''        End If
''    End If
''End If
''Set rst = Nothing
''
'''ADD INF TO REPORT
''
''Dim rss As Recordset
''
''''Core material
'''Set rst = New Recordset
'''rst.Open "select distinct core_mat from designspec where des_code = '" & DES & "'", con, adOpenStatic
'''If rst.EOF = False And rst.BOF = False Then
'''    CoreMat = rst!core_mat
'''End If
'''Set rst = Nothing
''
''Dim epoxy As Double
''Dim rstCRGO As Recordset
''
''Set rst = New Recordset
''rst.Open "select distinct core_no, tot_wt_core_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
'''    If CoreMat = "MU metal" Then
'''        '1***************
'''        If CoreShape = "Rectangluar" Then
'''            we = prcCalMUCoreShape("R", rst!tot_wt_core_wt)
'''        Else
'''            we = prcCalMUCoreShape("C", rst!tot_wt_core_wt)
'''        End If
'''    Else
''        '1***************
''
''        Do Until rst.EOF
''            Set rstCRGO = New Recordset
''            rstCRGO.Open "select core_mat, core_wt from designcore where dec_code = '" & DES & "' and coreno = " & rst!core_no & "", con, adOpenStatic, adLockReadOnly
''            If rstCRGO.EOF = False And rstCRGO.BOF = False Then CRGO = rstCRGO.Fields(0) & vbNullString
''
''            If CoreShape = "Rectangluar" Then
''                CoreID = FindCoreID(DES, "R", rst!core_no)
''                we = prcCalCoreShape("R", rstCRGO!core_wt, CoreID)
''            Else
''                CoreID = FindCoreID(DES, "C", rst!core_no)
''                we = prcCalCoreShape("C", rstCRGO!core_wt, CoreID)
''            End If
''            Set rstCRGO = Nothing
''            rst.MoveNext
''        Loop
'''    End If
''End If
''Set rst = Nothing
''
'''2************************
'''Super Enabled Copper
''Set rst = New Recordset
''rst.Open "select max(coreno) from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''    X = rst.Fields(0)
''End If
''Set rst = Nothing
''
''    'Finding swg of each core fro design core
''For i = 1 To X
''    Set rst = New Recordset
''    rst.Open "select swg, sec_cu_wt from designcore where dec_code = '" & DES & "' and coreno = " & i & " ", con, adOpenStatic
''    If rst.EOF = False And rst.BOF = False Then
''        y = RateCopper(rst!SWG, rst!sec_cu_wt)
''    End If
''    Set rst = Nothing
''Next
''
'''3***************
'''Copper bunch 1 or 2
''Set rst = New Recordset
''rst.Open "select distinct coreno from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''    vall = rst!coreno
''End If
''Set rst = Nothing
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''    '4**************
''    'Terminal
''    we = RateCal(CopTer1, 1)
''    we = RateCal(CopTer2, 1)
''
''
''
''    '' +++++++++++++
''    ''Copper strip tot_wt_pricu_wt
''    Dim rst2 As Recordset
''    Set rst2 = New Recordset
''    rst2.Open "select tot_wt_pricu_wt  from designspec where des_code = '" & DES & "'", con, adOpenStatic
''    If rst2.EOF = False And rst2.BOF = False Then
''        ''if only one group present all wt. taken by 1
''        If Bar2 = "" Then
''            we = RateCal(Bar1, CDbl(rst2!tot_wt_pricu_wt))
''        ''else wt. distributed
''        Else
''            we = RateCal(Bar1, (CDbl(rst2!tot_wt_pricu_wt) * 0.6))
''            we = RateCal(Bar2, (CDbl(rst2!tot_wt_pricu_wt) * 0.4))
''        End If
''    End If
''    Set rst2 = Nothing
''
''End If
''
''If InStr(1, UCase(Equip), "BAR") > 0 Then
''    we = RateCal(Bartype, 1)
''End If
''
'''++++++++++++++++++++
'''5 Resin
'''Dim epoxy As Double
''Set rst = New Recordset
''rst.Open "select distinct tot_wt_epoxy_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''     we = prcCalResin(CDbl(rst.Fields(0)))
''End If
''Set rst = Nothing
''
''
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''
''    '6**************
''    'Bracket
''    we = RateCal(Brackett, 1)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''Else
''
''    '6**************
''    'for insert
''    we = RateCal(CStr(WinBar), 4)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''End If
''
'''+++++++++++++++++++++++++++++++++
'''Rest of the installation
''Set rst1 = New Recordset
''rst1.Open "costing", con, adOpenDynamic, adLockOptimistic
''rst1.AddNew
''rst1!Item_name = "-"
''rst1.Update
''
''Dim rstItem As Recordset
''
''
''Set rst = New Recordset
''rst.Open "select * from new_costmaster where name = '" & EqType & "'", con, adOpenStatic, adLockReadOnly
'''rst.Open "select * from costingcomm", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''    Do Until rst.EOF
''        rst1.AddNew
''        rst1!code = rst!item_code
''        Set rstItem = New Recordset
''        rstItem.Open "select * from itemmast where item_code = '" & rst!item_code & "'", con, adOpenStatic, adLockReadOnly
''            rst1!Item_name = rstItem!item_desc1 & vbNullString
''            rst1!units = rstItem!uom & vbNullString
''            If IsNull(rstItem!landedprice) = True Then
''                rst1!Rate = 0
''            Else
''                rst1!Rate = Format(rstItem!landedprice, "0.00")
''            End If
''            rst1!qty = Format(Trim(rst!qty), "0.00")
''            rst1!amount = Format(CDbl(rst!qty) * CDbl(rst1!Rate), "0.00")
''        Set rstItem = Nothing
''        rst1.Update
''        rst.MoveNext
''    Loop
''    'Sundary Material
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Sundry Material"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Mould Charges
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Mould Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Packing Charges
''    rst1.AddNew
''    rst1!code = "----"
''    rst1!Item_name = "Packing Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''
''End If
''Set rst = Nothing
''
''Set rst1 = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''
''Private Function RateCal(item_code As String, CQty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
''Set rst = New Recordset
'''rst.Open "select * from costmaster where part = '" & Item_code & "'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from itemmast where item_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly
''
''If rst.EOF = False And rst.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!code = item_code
''      If InStr(rst!item_desc1 & vbNullString, "Insert") > 0 Then
''        If InStr(rst!item_desc1 & vbNullString, "M.S.") > 0 Then
''        'M.S.
''            If IsNull(rst!Type) = True Then
''                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!Mod & ")"
''            Else
''                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!Mod & ")" & " Type:" & rst!Type
''            End If
''        Else
''        'Brass
''            If IsNull(rst!Type) = True Then
''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")"
''            Else
''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''            End If
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "PLATE") > 0 Then
''      'Mounding Plate
''        If IsNull(rst!Type) = True Or Len(rst!Type) = 0 Then
''            If Len(Trim(rst!Mod)) = 0 Then
''                rs1!Item_name = "Mounting Plate " & rst!Drg
''            Else
''                rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")"
''            End If
''        Else
''            rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BRACKET") > 0 Then
''      'Mounting Bracket
''        If IsNull(rst!Type) = True And Len(Trim(rst!Type)) = 0 Then
''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")"
''        Else
''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "TERMINAL") > 0 Then
''      'Mounting Bracket
''        If IsNull(rst!Type) = True Then
''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")"
''        Else
''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BUNCHED") > 0 Then
''      'Mounting Bracket
''        rs1!Item_name = rst!item_desc2 & " group copper bunch"
''      Else
''      'Else
''        rs1!Item_name = rst!item_desc1
''        If Len(Trim(rst!item_desc4)) > 0 Or IsNull(rst!item_desc4) = False Then
''            rs1!Item_name = rs1!Item_name & " " & rst!item_desc4
''        End If
''
''      End If
''
''      rs1!units = rst!uom & vbNullString
''      rs1!Rate = Val(rst!landedprice & vbNullString)
''      rs1!qty = CQty
''      rs1!amount = Val(rst!landedprice & vbNullString) * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''
''End Function
''
''
''Private Function prcCalResin(qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For resin Only
''Set rst = New Recordset
'''rst.Open "select * from costmaster where part = 'INS-00A'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from itemmast where item_code = 'INS00A'", con, adOpenStatic, adLockReadOnly
''
''If rst.EOF = False And rst.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = "INS00A"
''     rs1!Item_name = "Hot Setting Epoxy Resin Max"
''     rs1!units = "Kgs"
''     rs1!Rate = rst!landedprice
''     rs1!qty = CInt(qty)
''     rs1!amount = rst!landedprice * CInt(qty)
''     rs1.Update
''     Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function prcCalMUCoreShape(TyprC As String, qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
''If TyprC = "R" Then
''    rrr = "CRGO Rectangular Maganetic Core"
''    rcode = "REC-00M"
''    RATEE = 2500
''Else
''    rrr = "CRGO Circular Maganetic Core"
''    rcode = "ROC-00M"
''    RATEE = 2500
''End If
''
''Dim rs1 As Recordset
''
''Set rst = New Recordset
''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''If rst.EOF = False And rst.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function GetCorePrice(CoreShape As String, Qtykg As Double) As Currency
''
''    Dim rs As New ADODB.Recordset
''    Dim SQL1 As String
''    Set rs = New ADODB.Recordset
''    If CoreShape = "R" Then
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRE00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRE00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRE00C"
''        End If
''    Else
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRO00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRO00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRO00C"
''        End If
''    End If
''
''
''    rs.Open "Select * from Core_rate where coretype like '" & CoreShape & "%' and grade = '" & Mid(CRGO, 6, 6) & "'", con, adOpenKeyset, adLockReadOnly
'''    rs.Open "select * from itemmast where item_code = '" & SQL1 & "'", con, adOpenStatic, adLockReadOnly
''    If rs.EOF = False And rs.BOF = False Then
''        GetCorePrice = rs!landedprice
''    End If
''    Set rs = Nothing
''End Function
''
''Private Function prcCalCoreShape(TyprC As String, qty As Double, CoreCode As String)
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
'' If TyprC = "R" Then
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("R", qty)
''Else
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("C", qty)
''End If
''
''Dim rs1 As Recordset
''
''Dim rst1 As ADODB.Recordset
''Set rst1 = New Recordset
'''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''rst1.Open "select * from ITEMMAST where item_code = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''
''If rst1.EOF = False And rst1.BOF = False Then
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''Else
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!code = "-"
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''
''Set rst1 = Nothing
''
''End Function
''
''Private Function RateCopper(item_code As String, CQty As Double)
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
''item_code = "SWG " & item_code
''Set rst = New Recordset
'''rst.Open "select * from costmaster where title like '" & Item_code & "%'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from itemmast where item_desc3 like '%" & item_code & "%'", con, adOpenStatic, adLockReadOnly
''
'''str1 = "select * from costmaster where title like '" & Item_code & "%'"
''If rst.EOF = False And rst.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!code = rst!item_code & vbNullString 'rst!part
''      rs1!Item_name = item_code & " Super Enamelled copper wire"
''      rs1!units = rst!uom & vbNullString
''      rs1!Rate = rst!landedprice
''      rs1!qty = CQty
''      rs1!amount = rst!landedprice * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set rst = Nothing
''
''End Function
''
''
''Dim CRGO As String
''Dim CoreSz As String
''Dim RST As Recordset
''
''Public Function FindCoreID(DesCode As String, CShape As String, CorNo As Integer) As String
''    Dim rs As New ADODB.Recordset
''    Dim rs1 As New ADODB.Recordset
''    Dim DWL, WWW, DOL, DOW, DHT As String
''    Dim DID, DOD As String
''    If CShape = "R" Then
''    'For Rectangular
''
''        Set rs = New ADODB.Recordset
''        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''        If rs.EOF = False And rs.BOF = False Then
'''            If rs!WW > rs!WL Then
''                DWw = CStr(rs!WW)
''                DWL = CStr(rs!WL)
''
''                DOL = CStr(rs!OL)
''                DOW = CStr(rs!OW)
''
''                DHT = CStr(rs!HT)
'''            Else
'''                DWL = CStr(rs!WW)
'''                DWW = CStr(rs!WL)
'''
'''                DOL = CStr(rs!OW)
'''                DOW = CStr(rs!OL)
'''
'''                DHT = CStr(rs!HT)
'''            End If
''            Dim y As Double
''            Dim X As Integer
''            If Len(DWL) = 4 Then
''                Size = DWL
''            ElseIf Len(DWL) = 3 Then
''                Size = DWL & "0"
''            ElseIf Len(DWL) = 2 Then
''                Size = "0" & DWL & "0"
''            ElseIf Len(DWL) = 1 Then
''                Size = "00" & DWL & "0"
''            End If
''
''            If Len(DWw) = 4 Then
''                Size = Size & DWw
''            ElseIf Len(DWw) = 3 Then
''                Size = Size & DWw & "0"
''            ElseIf Len(DWw) = 2 Then
''                Size = Size & "0" & DWw & "0"
''            ElseIf Len(DWw) = 1 Then
''                Size = Size & "00" & DWw & "0"
''            End If
''
''
''            y = (CDbl(DOL) - CDbl(DWL)) / 2
''            'txtOOd.Text = CDbl(DWW) + y * 2
''        '    MsgBox CInt(X)
''            If InStr(1, CStr(y), ".") > 0 Then
''                    X = y * 10
''                Else
''                    X = y
''            End If
''            If Len(CStr(X)) = 4 Then
''                Size = Size & X
''            ElseIf Len(CStr(X)) = 3 Then
''                Size = Size & "0" & X
''            ElseIf Len(CStr(X)) = 2 Then
''                Size = Size & "0" & X & "0"
''            ElseIf Len(CStr(X)) = 1 Then
''                Size = Size & "00" & X & "0"
''            End If
''
''            If Len(DHT) = 4 Then
''                Size = Size & DHT
''            ElseIf Len(DHT) = 3 Then
''                Size = Size & DHT & "0"
''            ElseIf Len(DHT) = 2 Then
''                Size = Size & "0" & DHT & "0"
''            ElseIf Len(DHT) = 1 Then
''                Size = Size & "00" & DHT & "0"
''            End If
''
'''            MsgBox Size
''            Dim RST As New ADODB.Recordset
''            Set RST = New ADODB.Recordset
''            RST.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DWL & "X" & DWw & " " & DOL & "X" & DOW & " " & DHT
''            If RST.EOF = False And RST.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''                FindCoreID = RST!item_code
''            Else
''                '"""""""
''                'reversing the size and checking again
''                DWL = CStr(rs!WW)
''                DWw = CStr(rs!WL)
''
''                DOL = CStr(rs!OW)
''                DOW = CStr(rs!OL)
''
''                DHT = CStr(rs!HT)
''                If Len(DWL) = 4 Then
''                    Size = DWL
''                ElseIf Len(DWL) = 3 Then
''                    Size = DWL & "0"
''                ElseIf Len(DWL) = 2 Then
''                    Size = "0" & DWL & "0"
''                ElseIf Len(DWL) = 1 Then
''                    Size = "00" & DWL & "0"
''                End If
''
''                If Len(DWw) = 4 Then
''                    Size = Size & DWw
''                ElseIf Len(DWw) = 3 Then
''                    Size = Size & DWw & "0"
''                ElseIf Len(DWw) = 2 Then
''                    Size = Size & "0" & DWw & "0"
''                ElseIf Len(DWw) = 1 Then
''                    Size = Size & "00" & DWw & "0"
''                End If
''
''
''                y = (CDbl(DOL) - CDbl(DWL)) / 2
''                'txtOOd.Text = CDbl(DWW) + y * 2
''            '    MsgBox CInt(X)
''                If InStr(1, CStr(y), ".") > 0 Then
''                        X = y * 10
''                    Else
''                        X = y
''                End If
''                If Len(CStr(X)) = 4 Then
''                    Size = Size & X
''                ElseIf Len(CStr(X)) = 3 Then
''                    Size = Size & "0" & X
''                ElseIf Len(CStr(X)) = 2 Then
''                    Size = Size & "0" & X & "0"
''                ElseIf Len(CStr(X)) = 1 Then
''                    Size = Size & "00" & X & "0"
''                End If
''
''                If Len(DHT) = 4 Then
''                    Size = Size & DHT
''                ElseIf Len(DHT) = 3 Then
''                    Size = Size & DHT & "0"
''                ElseIf Len(DHT) = 2 Then
''                    Size = Size & "0" & DHT & "0"
''                ElseIf Len(DHT) = 1 Then
''                    Size = Size & "00" & DHT & "0"
''                End If
''
'''                MsgBox Size
''                'Dim rst As New ADODB.Recordset
''                Set RST = New ADODB.Recordset
''                RST.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''                CoreSz = DWw & "X" & DWL & " " & DOW & "X" & DOL & " " & DHT
''                If RST.EOF = False And RST.BOF = False Then
'''                    MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''                    FindCoreID = RST!item_code
''                End If
''
''                '"""""""
''            End If
''
''        End If
''        'If found the Code
''
''
''        'Else generate new
''
''    Else
''    'For Circular
''        Set rs = New ADODB.Recordset
''        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''
''        If rs.EOF = False And rs.BOF = False Then
'''                If rs!ID > rs!od Then
''                    DID = CStr(rs!oD)
''                    DOD = CStr(rs!ID)
''                    DHT = CStr(rs!HT)
'''                Else
'''                    DID = CStr(rs!ID)
'''                    DOD = CStr(rs!od)
'''                    DHT = CStr(rs!HT)
'''                End If
''                Size = ""
''
''                If Len(DID) = 4 Then
''                    Size = DID
''                ElseIf Len(DID) = 3 Then
''                    Size = DID & "0"
''                ElseIf Len(DID) = 2 Then
''                    Size = "0" & DID & "0"
''                ElseIf Len(DID) = 1 Then
''                    Size = "00" & DID & "0"
''                End If
''
''                If Len(DOD) = 4 Then
''                    Size = Size & DOD
''                ElseIf Len(DOD) = 3 Then
''                    Size = Size & DOD & "0"
''                ElseIf Len(DOD) = 2 Then
''                    Size = Size & "0" & DOD & "0"
''                ElseIf Len(DOD) = 1 Then
''                    Size = Size & "00" & DOD & "0"
''                End If
''
''
''                If Len(DHT) = 3 Then
''                    Size = Size & DHT & "0"
''                ElseIf Len(DHT) = 2 Then
''                    Size = Size & "0" & DHT & "0"
''                ElseIf Len(DHT) = 1 Then
''                    Size = Size & "000" & DHT
''                End If
''
'''            MsgBox Size
'''            Dim rst As New ADODB.Recordset
''            Dim rst1 As ADODB.Recordset
''            Set rst1 = New ADODB.Recordset
''            rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DID & " X " & DOD & "  " & DHT
''            If rst1.EOF = False And rst1.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                FindCoreID = rst1!item_code
''            Else
''                'Reversing Size and checking again
''                    DID = CStr(rs!ID)
''                    DOD = CStr(rs!oD)
''                    DHT = CStr(rs!HT)
''                    Size = ""
''
''                    If Len(DID) = 4 Then
''                        Size = DID
''                    ElseIf Len(DID) = 3 Then
''                        Size = DID & "0"
''                    ElseIf Len(DID) = 2 Then
''                        Size = "0" & DID & "0"
''                    ElseIf Len(DID) = 1 Then
''                        Size = "00" & DID & "0"
''                    End If
''
''                    If Len(DOD) = 4 Then
''                        Size = Size & DOD
''                    ElseIf Len(DOD) = 3 Then
''                        Size = Size & DOD & "0"
''                    ElseIf Len(DOD) = 2 Then
''                        Size = Size & "0" & DOD & "0"
''                    ElseIf Len(DOD) = 1 Then
''                        Size = Size & "00" & DOD & "0"
''                    End If
''
''
''                    If Len(DHT) = 3 Then
''                        Size = Size & DHT & "0"
''                    ElseIf Len(DHT) = 2 Then
''                        Size = Size & "0" & DHT & "0"
''                    ElseIf Len(DHT) = 1 Then
''                        Size = Size & "000" & DHT
''                    End If
''
'''                MsgBox Size
''    '            Dim rst As New ADODB.Recordset
''                'Dim rst1 As ADODB.Recordset
''                Set rst1 = New ADODB.Recordset
''                rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''                CoreSz = DID & " X " & DOD & "  " & DHT
''                If rst1.EOF = False And rst1.BOF = False Then
'''                    MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                    FindCoreID = rst1!item_code
''                End If
''                'End Reversing Size and checking again
''            End If
''            Set rst1 = Nothing
''
''
''        End If
''    End If
''
''End Function
''
''Public Function Coating(DES As String)
''Dim EqType As String
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++
'''Deleteing coasting
''Set RST = New Recordset
''RST.Open "delete from costing", con, adOpenDynamic, adLockOptimistic
'' Set RST = Nothing
''
''Dim rst1 As Recordset
''
''Dim priCross As String
''Dim PTurns As Integer
''Dim CopTer1 As String
''Dim CopTer2 As String
''Dim Bar1 As String
''Dim Bar2 As String
''Dim Brackett As String
''Dim Plate As String
''Dim Drg As String
''Dim drgtype As String
''Dim Equip As String
''Dim CoreShape As String
''Dim CoreMat As String
''Dim CoreID As String
''Dim PriCond As String
''Dim PriCondType As String
''Dim PriConductor As String
''
''Dim Bartype As String
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Fing Pric Cross Sction
''Set RST = New Recordset
''RST.Open "select coreshape, pri_cond_style, equip, pri_cross_sec, pri_turns , drg_no , drg_type, pri_cond from designspec where des_code = '" & DES & "'", con, adOpenStatic, adLockReadOnly
''If RST.EOF = False And RST.BOF = False Then
''    priCross = RST!pri_cross_sec
''    PTurns = RST!pri_turns
''    Drg = RST!drg_no & vbNullString
''    drgtype = RST!DRG_TYPE & vbNullString
''    Equip = RST!Equip & vbNullString
''    CoreShape = RST!CoreShape & vbNullString
''    PriConductor = RST!pri_cond & vbNullString
''    If IsNull(RST!pri_cond_style) = False And Len(Trim(RST!pri_cond_style)) > 0 Then
''        PriCond = Mid(RST!pri_cond_style, 1, 1)
''        PriCondType = Mid(RST!pri_cond_style, InStr(1, RST!pri_cond_style, "-") + 1)
''    End If
''End If
''Set RST = Nothing
''
''If InStr(1, Equip, "Wound") > 0 Then
''    EqType = "2"
''ElseIf InStr(1, Equip, "Window") > 0 Then
''    EqType = "1"
''ElseIf InStr(1, Equip, "Bar") > 0 Then
''    EqType = "1"
''End If
''
''
''Dim Single_Or_Double As String
''If InStr(1, PriCondType, "Double") > 0 Then
''    Single_Or_Double = "Double"
''ElseIf InStr(1, PriCondType, "Single") > 0 Then
''    Single_Or_Double = "Single"
''ElseIf InStr(1, PriCondType, "Series") > 0 Then
''    Single_Or_Double = "Series"
''End If
''
''If Single_Or_Double = "Double" Then
''    If InStr(1, PriCondType, "Equal") > 0 Then
''        Single_Or_Double = "Double-Equal"
''    ElseIf InStr(1, PriCondType, "Different") > 0 Then
''        Single_Or_Double = "Double-Different"
''    Else
''        Single_Or_Double = "Double"
''    End If
''End If
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from primarycond11  -by providing c/s and turns
'''++++++++++++++++++++++++++++++
'''+changing ri primarycond from primartcond11
''Set RST = New Recordset
''str1 = "select * from primarycond11 where t_csp = '" & priCross & "' and t_np = '" & PTurns & "' and style = '" & Single_Or_Double & "'"
''If Trim(priCross) = "" Then
''    PriCond = 0
''    RST.Open "select *  from primarycond11 where [t_np] = " & PTurns & " and [type] = " & PriCond & " and [style] = '" & PriCondType & "' and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and [style] = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''Else
'''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and type = '" & PriCond & "%' and style = '" & PriCondType & "%' and drg = '" & Drg & "' and mod = '" & DrgType & "'", con, adOpenStatic, adLockReadOnly
''    strr = "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and mod = '" & drgtype & "' and style = '" & Single_Or_Double & "'"
''    RST.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and [mod] = '" & drgtype & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''End If
''If RST.EOF = False And RST.BOF = False Then
''    CopTer1 = RST!t_terp1 & vbNullString
''    CopTer2 = RST!t_terp2 & vbNullString
''    Bar1 = RST!cond_code1 & vbNullString
''    If Len(RST!cond_code2) > 0 And IsNull(RST!cond_code2) = False Then
''        Bar2 = RST!cond_code2
''    End If
''Else
''    Set RST = New Recordset
''    RST.Open "select *  from primarycond11 where t_np = " & PTurns & " and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and t_cond_siz = '" & PriConductor & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
''    If RST.EOF = False And RST.BOF = False Then
''        CopTer1 = RST!t_terp1 & vbNullString
''        CopTer2 = RST!t_terp2 & vbNullString
''        Bar1 = RST!cond_code1 & vbNullString
''        If Len(RST!cond_code2) > 0 Or IsNull(RST!cond_code2) = False Then
''            Bar2 = RST!cond_code2
''        End If
''    End If
''End If
''Set RST = Nothing
''
''
'''+++++++++++++++++++++++++++++++++++++++++++++
'''Finding code from MouldVol -by providing drg, type
''Set RST = New Recordset
''If Len(Trim(drgtype)) > 0 Then
''    RST.Open "select * from mouldvol2 where t_drg = '" & Drg & "' and t_drgtype = '" & drgtype & "'", con, adOpenStatic, adLockReadOnly
''Else
''    RST.Open "select * from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
''End If
''If RST.EOF = False And RST.BOF = False Then
''    If InStr(1, UCase(Equip), "WOUND") > 0 Then
''        Brackett = RST!t_bktcode & vbNullString  'for wound
''        Plate = RST!t_pltcode & vbNullString     'for wound
''    Else
''        WinBar = RST!t_mtginst & vbNullString   'for window and  Bar
''        Plate = RST!t_pltcode & vbNullString
''        If InStr(1, UCase(Equip), "BAR") > 0 Then
''            Bartype = RST!t_buscode & vbNullString
''        End If
''    End If
''End If
''Set RST = Nothing
''
'''ADD INF TO REPORT
''
''Dim rss As Recordset
''
''''Core material
'''Set rst = New Recordset
'''rst.Open "select distinct core_mat from designspec where des_code = '" & DES & "'", con, adOpenStatic
'''If rst.EOF = False And rst.BOF = False Then
'''    CoreMat = rst!core_mat
'''End If
'''Set rst = Nothing
''
''Dim epoxy As Double
''Dim rstCRGO As Recordset
''
''Set RST = New Recordset
''RST.Open "select distinct core_no, tot_wt_core_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If RST.EOF = False And RST.BOF = False Then
'''    If CoreMat = "MU metal" Then
'''        '1***************
'''        If CoreShape = "Rectangluar" Then
'''            we = prcCalMUCoreShape("R", rst!tot_wt_core_wt)
'''        Else
'''            we = prcCalMUCoreShape("C", rst!tot_wt_core_wt)
'''        End If
'''    Else
''        '1***************
''
''        Do Until RST.EOF
''            Set rstCRGO = New Recordset
''            rstCRGO.Open "select core_mat, core_wt from designcore where dec_code = '" & DES & "' and coreno = " & RST!core_no & "", con, adOpenStatic, adLockReadOnly
''            If rstCRGO.EOF = False And rstCRGO.BOF = False Then CRGO = rstCRGO.Fields(0) & vbNullString
''
''            If CoreShape = "Rectangluar" Then
''                CoreID = FindCoreID(DES, "R", RST!core_no)
''                we = prcCalCoreShape("R", rstCRGO!core_wt, CoreID)
''            Else
''                CoreID = FindCoreID(DES, "C", RST!core_no)
''                we = prcCalCoreShape("C", rstCRGO!core_wt, CoreID)
''            End If
''            Set rstCRGO = Nothing
''            RST.MoveNext
''        Loop
'''    End If
''End If
''Set RST = Nothing
''
'''2************************
'''Super Enabled Copper
''Set RST = New Recordset
''RST.Open "select max(coreno) from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If RST.EOF = False And RST.BOF = False Then
''    X = RST.Fields(0)
''End If
''Set RST = Nothing
''
''    'Finding swg of each core fro design core
''For i = 1 To X
''    Set RST = New Recordset
''    RST.Open "select swg, sec_cu_wt from designcore where dec_code = '" & DES & "' and coreno = " & i & " ", con, adOpenStatic
''    If RST.EOF = False And RST.BOF = False Then
''        y = RateCopper(RST!SWG, RST!sec_cu_wt)
''    End If
''    Set RST = Nothing
''Next
''
'''3***************
'''Copper bunch 1 or 2
''Set RST = New Recordset
''RST.Open "select distinct coreno from designcore where dec_code = '" & DES & "'", con, adOpenStatic
''If RST.EOF = False And RST.BOF = False Then
''    vall = RST!coreno
''End If
''Set RST = Nothing
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''    '4**************
''    'Terminal
''    we = RateCal(CopTer1, 1)
''    we = RateCal(CopTer2, 1)
''
''
''
''    '' +++++++++++++
''    ''Copper strip tot_wt_pricu_wt
''    Dim rst2 As Recordset
''    Set rst2 = New Recordset
''    rst2.Open "select tot_wt_pricu_wt  from designspec where des_code = '" & DES & "'", con, adOpenStatic
''    If rst2.EOF = False And rst2.BOF = False Then
''        ''if only one group present all wt. taken by 1
''        If Bar2 = "" Then
''            we = RateCal(Bar1, CDbl(rst2!tot_wt_pricu_wt))
''        ''else wt. distributed
''        Else
''            we = RateCal(Bar1, (CDbl(rst2!tot_wt_pricu_wt) * 0.6))
''            we = RateCal(Bar2, (CDbl(rst2!tot_wt_pricu_wt) * 0.4))
''        End If
''    End If
''    Set rst2 = Nothing
''
''End If
''
''If InStr(1, UCase(Equip), "BAR") > 0 Then
''    we = RateCal(Bartype, 1)
''End If
''
'''++++++++++++++++++++
'''5 Resin
'''Dim epoxy As Double
''Set RST = New Recordset
''RST.Open "select distinct tot_wt_epoxy_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If RST.EOF = False And RST.BOF = False Then
''     we = prcCalResin(CDbl(RST.Fields(0)))
''End If
''Set RST = Nothing
''
''
''
''If InStr(1, UCase(Equip), "WOUND") > 0 Then
''
''    '6**************
''    'Bracket
''    we = RateCal(Brackett, 1)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''Else
''
''    '6**************
''    'for insert
''    we = RateCal(CStr(WinBar), 4)
''
''    '7**************
''    'Plate
''    we = RateCal(Plate, 1)
''
''End If
''
'''+++++++++++++++++++++++++++++++++
'''Rest of the installation
''Set rst1 = New Recordset
''rst1.Open "costing", con, adOpenDynamic, adLockOptimistic
''rst1.AddNew
''rst1!Item_name = "-"
''rst1.Update
''
''Dim rstItem As Recordset
''
''
''Set RST = New Recordset
''RST.Open "select * from new_costmaster where name = '" & EqType & "'", con, adOpenStatic, adLockReadOnly
'''rst.Open "select * from costingcomm", con, adOpenStatic, adLockReadOnly
''If RST.EOF = False And RST.BOF = False Then
''    Do Until RST.EOF
''        rst1.AddNew
''        rst1!CODE = RST!item_code
''        Set rstItem = New Recordset
''        rstItem.Open "select * from itemmast where item_code = '" & RST!item_code & "'", con, adOpenStatic, adLockReadOnly
''            rst1!Item_name = rstItem!item_desc1 & vbNullString
''            rst1!units = rstItem!uom & vbNullString
''            If IsNull(rstItem!landedprice) = True Then
''                rst1!Rate = 0
''            Else
''                rst1!Rate = Format(rstItem!landedprice, "0.00")
''            End If
''            rst1!qty = Format(Trim(RST!qty), "0.00")
''            rst1!amount = Format(CDbl(RST!qty) * CDbl(rst1!Rate), "0.00")
''        Set rstItem = Nothing
''        rst1.Update
''        RST.MoveNext
''    Loop
''    'Sundary Material
''    rst1.AddNew
''    rst1!CODE = "----"
''    rst1!Item_name = "Sundry Material"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Mould Charges
''    rst1.AddNew
''    rst1!CODE = "----"
''    rst1!Item_name = "Mould Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''    'Packing Charges
''    rst1.AddNew
''    rst1!CODE = "----"
''    rst1!Item_name = "Packing Charges"
''    rst1!units = "Nos"
''    rst1!Rate = 100
''    rst1!qty = 1
''    rst1!amount = 100
''    rst1.Update
''
''End If
''Set RST = Nothing
''
''Set rst1 = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''
''Private Function RateCal(item_code As String, CQty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
''Set RST = New Recordset
'''rst.Open "select * from costmaster where part = '" & Item_code & "'", con, adOpenStatic, adLockReadOnly
''RST.Open "select * from itemmast where item_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly
''
''If RST.EOF = False And RST.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!CODE = item_code
''      If InStr(RST!item_desc1 & vbNullString, "Insert") > 0 Then
''        If InStr(RST!item_desc1 & vbNullString, "M.S.") > 0 Then
''        'M.S.
''            If IsNull(RST!Type) = True Then
''                rs1!Item_name = "M.S. Insert " & RST!Drg & "(" & RST!Mod & ")"
''            Else
''                rs1!Item_name = "M.S. Insert " & RST!Drg & "(" & RST!Mod & ")" & " Type:" & RST!Type
''            End If
''        Else
''        'Brass
''            If IsNull(RST!Type) = True Then
''                rs1!Item_name = "Brass Insert " & RST!Drg & "(" & RST!Mod & ")"
''            Else
''                rs1!Item_name = "Brass Insert " & RST!Drg & "(" & RST!Mod & ")" & "Type:" & RST!Type
''            End If
''        End If
''      ElseIf InStr(UCase(RST!item_desc1 & vbNullString), "PLATE") > 0 Then
''      'Mounding Plate
''        If IsNull(RST!Type) = True Or Len(RST!Type) = 0 Then
''            If Len(Trim(RST!Mod)) = 0 Then
''                rs1!Item_name = "Mounting Plate " & RST!Drg
''            Else
''                rs1!Item_name = "Mounting Plate " & RST!Drg & "(" & RST!Mod & ")"
''            End If
''        Else
''            rs1!Item_name = "Mounting Plate " & RST!Drg & "(" & RST!Mod & ")" & "Type:" & RST!Type
''        End If
''      ElseIf InStr(UCase(RST!item_desc1 & vbNullString), "BRACKET") > 0 Then
''      'Mounting Bracket
''        If IsNull(RST!Type) = True And Len(Trim(RST!Type)) = 0 Then
''            rs1!Item_name = "Mounting Bracket " & RST!Drg & "(" & RST!Mod & ")"
''        Else
''            rs1!Item_name = "Mounting Bracket " & RST!Drg & "(" & RST!Mod & ")" & "Type:" & RST!Type
''        End If
''      ElseIf InStr(UCase(RST!item_desc1 & vbNullString), "TERMINAL") > 0 Then
''      'Mounting Bracket
''        If IsNull(RST!Type) = True Then
''            rs1!Item_name = "Copper Terminal " & RST!Drg & "(" & RST!Mod & ")"
''        Else
''            rs1!Item_name = "Copper Terminal " & RST!Drg & "(" & RST!Mod & ")" & "Type:" & RST!Type
''        End If
''      ElseIf InStr(UCase(RST!item_desc1 & vbNullString), "BUNCHED") > 0 Then
''      'Mounting Bracket
''        rs1!Item_name = RST!item_desc2 & " group copper bunch"
''      Else
''      'Else
''        rs1!Item_name = RST!item_desc1
''        If Len(Trim(RST!item_desc4)) > 0 Or IsNull(RST!item_desc4) = False Then
''            rs1!Item_name = rs1!Item_name & " " & RST!item_desc4
''        End If
''
''      End If
''
''      rs1!units = RST!uom & vbNullString
''      rs1!Rate = Val(RST!landedprice & vbNullString)
''      rs1!qty = CQty
''      rs1!amount = Val(RST!landedprice & vbNullString) * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set RST = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''
''End Function
''
''
''Private Function prcCalResin(qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For resin Only
''Set RST = New Recordset
'''rst.Open "select * from costmaster where part = 'INS-00A'", con, adOpenStatic, adLockReadOnly
''RST.Open "select * from itemmast where item_code = 'INS00A'", con, adOpenStatic, adLockReadOnly
''
''If RST.EOF = False And RST.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!CODE = "INS00A"
''     rs1!Item_name = "Hot Setting Epoxy Resin Max"
''     rs1!units = "Kgs"
''     rs1!Rate = RST!landedprice
''     rs1!qty = CInt(qty)
''     rs1!amount = RST!landedprice * CInt(qty)
''     rs1.Update
''     Set rs1 = Nothing
''
''End If
''Set RST = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function prcCalMUCoreShape(TyprC As String, qty As Double)
''
''On Error GoTo err
''
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
''If TyprC = "R" Then
''    rrr = "CRGO Rectangular Maganetic Core"
''    rcode = "REC-00M"
''    RATEE = 2500
''Else
''    rrr = "CRGO Circular Maganetic Core"
''    rcode = "ROC-00M"
''    RATEE = 2500
''End If
''
''Dim rs1 As Recordset
''
''Set RST = New Recordset
''RST.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''If RST.EOF = False And RST.BOF = False Then
''
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!CODE = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''Set RST = Nothing
''
''err:
''    If Len(Trim(err.Description)) > 0 Then
''        MsgBox err.Description & vbCrLf & err.Number, vbCritical
''    End If
''
''End Function
''
''Private Function GetCorePrice(CoreShape As String, Qtykg As Double) As Currency
''
''    Dim rs As New ADODB.Recordset
''    Dim SQL1 As String
''    Set rs = New ADODB.Recordset
''    If CoreShape = "R" Then
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRE00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRE00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRE00C"
''        End If
''    Else
''        If Mid(CRGO, 6, 6) = "A" Then
''            SQL1 = "CRO00A"
''        ElseIf Mid(CRGO, 6, 6) = "B" Then
''            SQL1 = "CRO00B"
''        ElseIf Mid(CRGO, 6, 6) = "C" Then
''            SQL1 = "CRO00C"
''        End If
''    End If
''
''
''    rs.Open "Select * from Core_rate where coretype like '" & CoreShape & "%' and grade = '" & Mid(CRGO, 6, 6) & "'", con, adOpenKeyset, adLockReadOnly
'''    rs.Open "select * from itemmast where item_code = '" & SQL1 & "'", con, adOpenStatic, adLockReadOnly
''    If rs.EOF = False And rs.BOF = False Then
''        GetCorePrice = rs!landedprice
''    End If
''    Set rs = Nothing
''End Function
''
''Private Function prcCalCoreShape(TyprC As String, qty As Double, CoreCode As String)
'''++++++++++++++++++++++++++++++++++++++
'''For Core Only
''
'' If TyprC = "R" Then
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("R", qty)
''Else
''    rrr = CoreSz & " " & CRGO
''    rcode = CoreCode
''    RATEE = GetCorePrice("C", qty)
''End If
''
''Dim rs1 As Recordset
''
''Dim rst1 As ADODB.Recordset
''Set rst1 = New Recordset
'''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''rst1.Open "select * from ITEMMAST where item_code = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
''
''If rst1.EOF = False And rst1.BOF = False Then
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!CODE = rcode
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''Else
''     Set rs1 = New Recordset
''     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''
''     rs1.AddNew
''     rs1!CODE = "-"
''     rs1!Item_name = rrr
''     rs1!units = "Kgs"
''     rs1!Rate = RATEE
''     rs1!qty = qty
''     rs1!amount = qty * RATEE
''     rs1.Update
''
''     Set rs1 = Nothing
''
''End If
''
''Set rst1 = Nothing
''
''End Function
''
''Private Function RateCopper(item_code As String, CQty As Double)
'''++++++++++++++++++++++++++++++++++++++
'''Taking rate from costmaster
''
''Dim rs1 As Recordset
''
''item_code = "SWG " & item_code
''Set RST = New Recordset
'''rst.Open "select * from costmaster where title like '" & Item_code & "%'", con, adOpenStatic, adLockReadOnly
''RST.Open "select * from itemmast where item_desc3 like '%" & item_code & "%'", con, adOpenStatic, adLockReadOnly
''
'''str1 = "select * from costmaster where title like '" & Item_code & "%'"
''If RST.EOF = False And RST.BOF = False Then
''
''      Set rs1 = New Recordset
''      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
''      rs1.AddNew
''      rs1!CODE = RST!item_code & vbNullString 'rst!part
''      rs1!Item_name = item_code & " Super Enamelled copper wire"
''      rs1!units = RST!uom & vbNullString
''      rs1!Rate = RST!landedprice
''      rs1!qty = CQty
''      rs1!amount = RST!landedprice * CQty
''      rs1.Update
''      Set rs1 = Nothing
''
''End If
''Set RST = Nothing
''
''End Function
''
'Dim CRGO As String
'Dim CoreSz As String
'Dim rst As Recordset
'
'Public Function FindCoreID(DesCode As String, CShape As String, CorNo As Integer) As String
'
'    Dim rs As New ADODB.Recordset
'    Dim rs1 As New ADODB.Recordset
'    Dim DWL, WWW, DOL, DOW, DHT As String
'    Dim DID, DOD As String
'    If CShape = "R" Then
'    'For Rectangular
'        Set rs = New ADODB.Recordset
'        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
'        If rs.EOF = False And rs.BOF = False Then
'
'            DWw = CStr(rs!WW)
'            DWL = CStr(rs!WL)
'            DOL = CStr(rs!OL)
'            DOW = CStr(rs!OW)
'            DHT = CStr(rs!HT)
'
'            Dim y As Double
'            Dim X As Integer
'
'            Size = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
'
''            MsgBox Size
'            Dim rst As New ADODB.Recordset
'            Set rst = New ADODB.Recordset
'            rst.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'RECTANGULAR' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'            CoreSz = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
'            If rst.EOF = False And rst.BOF = False Then
'                FindCoreID = rst!product_code
'            End If
'
'        End If
'    Else
'    'For Circular
'        Set rs = New ADODB.Recordset
'        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
'        If rs.EOF = False And rs.BOF = False Then
'            DID = CStr(rs!oD)
'            DOD = CStr(rs!ID)
'            DHT = CStr(rs!HT)
'            Size = DOD & "X" & DID & "X" & DHT
'
'            Dim rst1 As ADODB.Recordset
'            Set rst1 = New ADODB.Recordset
'
'            rst1.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'ROUND;' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'            CoreSz = DOD & "X" & DID & "X" & DHT
'            If rst1.EOF = False And rst1.BOF = False Then
'                FindCoreID = rst1!product_code
'            End If
'            Set rst1 = Nothing
'
'        End If
'    End If
'    Set srt = Nothing
'Exit Function
''    Dim rs As New ADODB.Recordset
''    Dim rs1 As New ADODB.Recordset
''    Dim DWL, WWW, DOL, DOW, DHT As String
''    Dim DID, DOD As String
''    If CShape = "R" Then
''    'For Rectangular
''
''        Set rs = New ADODB.Recordset
''        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''        If rs.EOF = False And rs.BOF = False Then
'''            If rs!WW > rs!WL Then
''                DWw = CStr(rs!WW)
''                DWL = CStr(rs!WL)
''
''                DOL = CStr(rs!OL)
''                DOW = CStr(rs!OW)
''
''                DHT = CStr(rs!HT)
'''            Else
'''                DWL = CStr(rs!WW)
'''                DWW = CStr(rs!WL)
'''
'''                DOL = CStr(rs!OW)
'''                DOW = CStr(rs!OL)
'''
'''                DHT = CStr(rs!HT)
'''            End If
''            '+++++++++++++++++++++++++Pragati++++++++++++++
''            Dim y As Double
''            Dim X As Integer
'''            If Len(DWL) = 4 Then
'''                Size = DWL
'''            ElseIf Len(DWL) = 3 Then
'''                Size = DWL & "0"
'''            ElseIf Len(DWL) = 2 Then
'''                Size = "0" & DWL & "0"
'''            ElseIf Len(DWL) = 1 Then
'''                Size = "00" & DWL & "0"
'''            End If
'''            If Len(DWw) = 4 Then
'''                Size = Size & DWw
'''            ElseIf Len(DWw) = 3 Then
'''                Size = Size & DWw & "0"
'''            ElseIf Len(DWw) = 2 Then
'''                Size = Size & "0" & DWw & "0"
'''            ElseIf Len(DWw) = 1 Then
'''                Size = Size & "00" & DWw & "0"
'''            End If
'''            y = (CDbl(DOL) - CDbl(DWL)) / 2
'''            'txtOOd.Text = CDbl(DWW) + y * 2
'''        '    MsgBox CInt(X)
'''            If InStr(1, CStr(y), ".") > 0 Then
'''                    X = y * 10
'''                Else
'''                    X = y
'''            End If
'''            If Len(CStr(X)) = 4 Then
'''                Size = Size & X
'''            ElseIf Len(CStr(X)) = 3 Then
'''                Size = Size & "0" & X
'''            ElseIf Len(CStr(X)) = 2 Then
'''                Size = Size & "0" & X & "0"
'''            ElseIf Len(CStr(X)) = 1 Then
'''                Size = Size & "00" & X & "0"
'''            End If
'''
'''            If Len(DHT) = 4 Then
'''                Size = Size & DHT
'''            ElseIf Len(DHT) = 3 Then
'''                Size = Size & DHT & "0"
'''            ElseIf Len(DHT) = 2 Then
'''                Size = Size & "0" & DHT & "0"
'''            ElseIf Len(DHT) = 1 Then
'''                Size = Size & "00" & DHT & "0"
'''            End If
''
''            '++++++++++++++++++++++++++++++++++++++ Pragati end +++++++++++++++
''
''            Size = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
''
'''            MsgBox Size
''            Dim rst As New ADODB.Recordset
''            Set rst = New ADODB.Recordset
''            rst.Open "Select Product_code from MST_PART where Material_Type = 'TORROIDAL CRGO SILICON STEEL CORE;' and Other_Det_Type = 'ROUND;' and Dimension = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = "(" & DWL & "X" & DWw & ")*(" & DOL & "X" & DOW & ")*" & DHT
''            If rst.EOF = False And rst.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst!Item_code
''                FindCoreID = rst!Product_code
'''            Else
''
''                '"""""""
''                'reversing the size and checking again
'''                DWL = CStr(rs!WW)
'''                DWw = CStr(rs!WL)
'''
'''                DOL = CStr(rs!OW)
'''                DOW = CStr(rs!OL)
'''
'''                DHT = CStr(rs!HT)
'''                If Len(DWL) = 4 Then
'''                    Size = DWL
'''                ElseIf Len(DWL) = 3 Then
'''                    Size = DWL & "0"
'''                ElseIf Len(DWL) = 2 Then
'''                    Size = "0" & DWL & "0"
'''                ElseIf Len(DWL) = 1 Then
'''                    Size = "00" & DWL & "0"
'''                End If
'''
'''                If Len(DWw) = 4 Then
'''                    Size = Size & DWw
'''                ElseIf Len(DWw) = 3 Then
'''                    Size = Size & DWw & "0"
'''                ElseIf Len(DWw) = 2 Then
'''                    Size = Size & "0" & DWw & "0"
'''                ElseIf Len(DWw) = 1 Then
'''                    Size = Size & "00" & DWw & "0"
'''                End If
'''
'''
'''                y = (CDbl(DOL) - CDbl(DWL)) / 2
''                'txtOOd.Text = CDbl(DWW) + y * 2
''            '    MsgBox CInt(X)
'''                If InStr(1, CStr(y), ".") > 0 Then
'''                        X = y * 10
'''                    Else
'''                        X = y
'''                End If
'''                If Len(CStr(X)) = 4 Then
'''                    Size = Size & X
'''                ElseIf Len(CStr(X)) = 3 Then
'''                    Size = Size & "0" & X
'''                ElseIf Len(CStr(X)) = 2 Then
'''                    Size = Size & "0" & X & "0"
'''                ElseIf Len(CStr(X)) = 1 Then
'''                    Size = Size & "00" & X & "0"
'''                End If
'''
'''                If Len(DHT) = 4 Then
'''                    Size = Size & DHT
'''                ElseIf Len(DHT) = 3 Then
'''                    Size = Size & DHT & "0"
'''                ElseIf Len(DHT) = 2 Then
'''                    Size = Size & "0" & DHT & "0"
'''                ElseIf Len(DHT) = 1 Then
'''                    Size = Size & "00" & DHT & "0"
'''                End If
'''
''''                MsgBox Size
'''                'Dim rst As New ADODB.Recordset
'''                Set rst = New ADODB.Recordset
'''                rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
'''                CoreSz = DWw & "X" & DWL & " " & DOW & "X" & DOL & " " & DHT
'''                If rst.EOF = False And rst.BOF = False Then
''''                    MsgBox "The item you are searching for is : " & "  " & rst!Item_code
'''                    FindCoreID = rst!item_code
'''                End If
''
''                '"""""""
''            End If
''
''        End If
''        'If found the Code
''
''
''        'Else generate new
''
''    Else
''    'For Circular
''        Set rs = New ADODB.Recordset
''        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
''
''        If rs.EOF = False And rs.BOF = False Then
'''                If rs!ID > rs!od Then
''                    DID = CStr(rs!oD)
''                    DOD = CStr(rs!ID)
''                    DHT = CStr(rs!HT)
'''                Else
'''                    DID = CStr(rs!ID)
'''                    DOD = CStr(rs!od)
'''                    DHT = CStr(rs!HT)
'''                End If
''                Size = ""
''
''
''                If Len(DID) = 4 Then
''                    Size = DID
''                ElseIf Len(DID) = 3 Then
''                    Size = DID & "0"
''                ElseIf Len(DID) = 2 Then
''                    Size = "0" & DID & "0"
''                ElseIf Len(DID) = 1 Then
''                    Size = "00" & DID & "0"
''                End If
''
''                If Len(DOD) = 4 Then
''                    Size = Size & DOD
''                ElseIf Len(DOD) = 3 Then
''                    Size = Size & DOD & "0"
''                ElseIf Len(DOD) = 2 Then
''                    Size = Size & "0" & DOD & "0"
''                ElseIf Len(DOD) = 1 Then
''                    Size = Size & "00" & DOD & "0"
''                End If
''
''
''                If Len(DHT) = 3 Then
''                    Size = Size & DHT & "0"
''                ElseIf Len(DHT) = 2 Then
''                    Size = Size & "0" & DHT & "0"
''                ElseIf Len(DHT) = 1 Then
''                    Size = Size & "000" & DHT
''                End If
''
'''            MsgBox Size
'''            Dim rst As New ADODB.Recordset
''            Dim rst1 As ADODB.Recordset
''            Set rst1 = New ADODB.Recordset
''            rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''            CoreSz = DID & " X " & DOD & "  " & DHT
''            If rst1.EOF = False And rst1.BOF = False Then
'''                MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                FindCoreID = rst1!item_code
''            Else
''                'Reversing Size and checking again
''                    DID = CStr(rs!ID)
''                    DOD = CStr(rs!oD)
''                    DHT = CStr(rs!HT)
''                    Size = ""
''
''                    If Len(DID) = 4 Then
''                        Size = DID
''                    ElseIf Len(DID) = 3 Then
''                        Size = DID & "0"
''                    ElseIf Len(DID) = 2 Then
''                        Size = "0" & DID & "0"
''                    ElseIf Len(DID) = 1 Then
''                        Size = "00" & DID & "0"
''                    End If
''
''                    If Len(DOD) = 4 Then
''                        Size = Size & DOD
''                    ElseIf Len(DOD) = 3 Then
''                        Size = Size & DOD & "0"
''                    ElseIf Len(DOD) = 2 Then
''                        Size = Size & "0" & DOD & "0"
''                    ElseIf Len(DOD) = 1 Then
''                        Size = Size & "00" & DOD & "0"
''                    End If
''
''
''                    If Len(DHT) = 3 Then
''                        Size = Size & DHT & "0"
''                    ElseIf Len(DHT) = 2 Then
''                        Size = Size & "0" & DHT & "0"
''                    ElseIf Len(DHT) = 1 Then
''                        Size = Size & "000" & DHT
''                    End If
''
'''                MsgBox Size
''    '            Dim rst As New ADODB.Recordset
''                'Dim rst1 As ADODB.Recordset
''                Set rst1 = New ADODB.Recordset
''                rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
''                CoreSz = DID & " X " & DOD & "  " & DHT
''                If rst1.EOF = False And rst1.BOF = False Then
'''                    MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
''                    FindCoreID = rst1!item_code
''                End If
''                'End Reversing Size and checking again
''            End If
''            Set rst1 = Nothing
''
''
''        End If
''    End If
'
'End Function
'
'Public Function Coating(DES As String)
'Dim EqType As String
'
'On Error GoTo err
'
''++++++++++++++++++++++++++++++++++++
''Deleteing coasting
'Set rst = New Recordset
'rst.Open "delete from costing", con, adOpenDynamic, adLockOptimistic
' Set rst = Nothing
'
'Dim rst1 As Recordset
'
'Dim priCross As String
'Dim PTurns As Integer
'Dim CopTer1 As String
'Dim CopTer2 As String
'Dim Bar1 As String
'Dim Bar2 As String
'Dim Brackett As String
'Dim Plate As String
'Dim Drg As String
'Dim drgtype As String
'Dim Equip As String
'Dim CoreShape As String
'Dim CoreMat As String
'Dim CoreID As String
'Dim PriCond As String
'Dim PriCondType As String
'Dim PriConductor As String
'
'Dim Bartype As String
'
''+++++++++++++++++++++++++++++++++++++++++++++
''Fing Pric Cross Sction
'Set rst = New Recordset
'rst.Open "select coreshape, pri_cond_style, equip, pri_cross_sec, pri_turns , drg_no , drg_type, pri_cond from designspec where des_code = '" & DES & "'", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'    priCross = rst!pri_cross_sec
'    PTurns = rst!pri_turns
'    Drg = rst!drg_no & vbNullString
'    drgtype = rst!DRG_TYPE & vbNullString
'    Equip = rst!Equip & vbNullString
'    CoreShape = rst!CoreShape & vbNullString
'    PriConductor = rst!pri_cond & vbNullString
'    'Only primary condutor involved
'    If IsNull(rst!pri_cond_style) = False And Len(Trim(rst!pri_cond_style)) > 0 Then
'        PriCond = Mid(rst!pri_cond_style, 1, 1)
'        PriCondType = Mid(rst!pri_cond_style, InStr(1, rst!pri_cond_style, "-") + 1)
'    End If
'End If
'Set rst = Nothing
'
'If InStr(1, Equip, "Wound") > 0 Then
'    EqType = "2"
'ElseIf InStr(1, Equip, "Window") > 0 Then
'    EqType = "1"
'ElseIf InStr(1, Equip, "Bar") > 0 Then
'    EqType = "1"
'End If
'
'
'Dim Single_Or_Double As String
'If InStr(1, PriCondType, "Double") > 0 Then
'    Single_Or_Double = "Double"
'ElseIf InStr(1, PriCondType, "Single") > 0 Then
'    Single_Or_Double = "Single"
'ElseIf InStr(1, PriCondType, "Series") > 0 Then
'    Single_Or_Double = "Series"
'End If
'
'If Single_Or_Double = "Double" Then
'    If InStr(1, PriCondType, "Equal") > 0 Then
'        Single_Or_Double = "Double-Equal"
'    ElseIf InStr(1, PriCondType, "Different") > 0 Then
'        Single_Or_Double = "Double-Different"
'    Else
'        Single_Or_Double = "Double"
'    End If
'End If
'
''+++++++++++++++++++++++++++++++++++++++++++++
''Finding code from primarycond11  -by providing c/s and turns
''++++++++++++++++++++++++++++++
''+changing ri primarycond from primartcond11
'Set rst = New Recordset
'str1 = "select * from primarycond11 where t_csp = '" & priCross & "' and t_np = '" & PTurns & "' and style = '" & Single_Or_Double & "'"
'If Trim(priCross) = "" Then
'    PriCond = 0
'    rst.Open "select *  from primarycond11 where [t_np] = " & PTurns & " and [type] = " & PriCond & " and [style] = '" & PriCondType & "' and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and [style] = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
'Else
''    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and type = '" & PriCond & "%' and style = '" & PriCondType & "%' and drg = '" & Drg & "' and mod = '" & DrgType & "'", con, adOpenStatic, adLockReadOnly
'    strr = "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and mod = '" & drgtype & "' and style = '" & Single_Or_Double & "'"
'    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and [mod] = '" & drgtype & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
'End If
'If rst.EOF = False And rst.BOF = False Then
'    CopTer1 = rst!t_terp1 & vbNullString
'    CopTer2 = rst!t_terp2 & vbNullString
'    Bar1 = rst!cond_code1 & vbNullString
'    If Len(rst!cond_code2) > 0 And IsNull(rst!cond_code2) = False Then
'        Bar2 = rst!cond_code2
'    End If
'Else
'    Set rst = New Recordset
'    rst.Open "select *  from primarycond11 where t_np = " & PTurns & " and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and t_cond_siz = '" & PriConductor & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
'    If rst.EOF = False And rst.BOF = False Then
'        CopTer1 = rst!t_terp1 & vbNullString
'        CopTer2 = rst!t_terp2 & vbNullString
'        Bar1 = rst!cond_code1 & vbNullString
'        If Len(rst!cond_code2) > 0 Or IsNull(rst!cond_code2) = False Then
'            Bar2 = rst!cond_code2
'        End If
'    End If
'End If
'Set rst = Nothing
'
'
''+++++++++++++++++++++++++++++++++++++++++++++
''Finding code from MouldVol -by providing drg, type
'Set rst = New Recordset
'If Len(Trim(drgtype)) > 0 Then
'    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "' and t_drgtype = '" & drgtype & "'", con, adOpenStatic, adLockReadOnly
'Else
'    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
'End If
'If rst.EOF = False And rst.BOF = False Then
'    If InStr(1, UCase(Equip), "WOUND") > 0 Then
'        Brackett = rst!t_bktcode & vbNullString  'for wound
'        Plate = rst!t_pltcode & vbNullString     'for wound
'    Else
'        WinBar = rst!t_mtginst & vbNullString   'for window and  Bar
'        Plate = rst!t_pltcode & vbNullString
'        If InStr(1, UCase(Equip), "BAR") > 0 Then
'            Bartype = rst!t_buscode & vbNullString
'        End If
'    End If
'End If
'Set rst = Nothing
'
''ADD INF TO REPORT
'
'Dim rss As Recordset
'
'''Core material
''Set rst = New Recordset
''rst.Open "select distinct core_mat from designspec where des_code = '" & DES & "'", con, adOpenStatic
''If rst.EOF = False And rst.BOF = False Then
''    CoreMat = rst!core_mat
''End If
''Set rst = Nothing
'
'Dim epoxy As Double
'Dim rstCRGO As Recordset
'
'Set rst = New Recordset
'rst.Open "select distinct core_no, tot_wt_core_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
'If rst.EOF = False And rst.BOF = False Then
''    If CoreMat = "MU metal" Then
''        '1***************
''        If CoreShape = "Rectangluar" Then
''            we = prcCalMUCoreShape("R", rst!tot_wt_core_wt)
''        Else
''            we = prcCalMUCoreShape("C", rst!tot_wt_core_wt)
''        End If
''    Else
'        '1***************
'
'        Do Until rst.EOF
'            Set rstCRGO = New Recordset
'            rstCRGO.Open "select core_mat, core_wt from designcore where dec_code = '" & DES & "' and coreno = " & rst!core_no & "", con, adOpenStatic, adLockReadOnly
'            If rstCRGO.EOF = False And rstCRGO.BOF = False Then CRGO = rstCRGO.Fields(0) & vbNullString
'
'            If CoreShape = "Rectangluar" Then
'                CoreID = FindCoreID(DES, "R", rst!core_no)
'                we = prcCalCoreShape("R", rstCRGO!core_wt, CoreID)
'            Else
'                CoreID = FindCoreID(DES, "C", rst!core_no)
'                we = prcCalCoreShape("C", rstCRGO!core_wt, CoreID)
'            End If
'            Set rstCRGO = Nothing
'            rst.MoveNext
'        Loop
''    End If
'End If
'Set rst = Nothing
'
''2************************
''Super Enabled Copper
'Set rst = New Recordset
'rst.Open "select max(coreno) from designcore where dec_code = '" & DES & "'", con, adOpenStatic
'If rst.EOF = False And rst.BOF = False Then
'    X = rst.Fields(0)
'End If
'Set rst = Nothing
'
'    'Finding swg of each core fro design core
'For i = 1 To X
'    Set rst = New Recordset
'    rst.Open "select swg, sec_cu_wt from designcore where dec_code = '" & DES & "' and coreno = " & i & " ", con, adOpenStatic
'    If rst.EOF = False And rst.BOF = False Then
'        y = RateCopper(rst!SWG, rst!sec_cu_wt)
'    End If
'    Set rst = Nothing
'Next
'
''3***************
''Copper bunch 1 or 2
'Set rst = New Recordset
'rst.Open "select distinct coreno from designcore where dec_code = '" & DES & "'", con, adOpenStatic
'If rst.EOF = False And rst.BOF = False Then
'    vall = rst!coreno
'End If
'Set rst = Nothing
'
'If InStr(1, UCase(Equip), "WOUND") > 0 Then
'    '4**************
'    'Terminal
'    we = RateCal(CopTer1, 1)
'    we = RateCal(CopTer2, 1)
'
'
'
'    '' +++++++++++++
'    ''Copper strip tot_wt_pricu_wt
'    Dim rst2 As Recordset
'    Set rst2 = New Recordset
'    rst2.Open "select tot_wt_pricu_wt  from designspec where des_code = '" & DES & "'", con, adOpenStatic
'    If rst2.EOF = False And rst2.BOF = False Then
'        ''if only one group present all wt. taken by 1
'        If Bar2 = "" Then
'            we = RateCal(Bar1, CDbl(rst2!tot_wt_pricu_wt))
'        ''else wt. distributed
'        Else
'            we = RateCal(Bar1, (CDbl(rst2!tot_wt_pricu_wt) * 0.6))
'            we = RateCal(Bar2, (CDbl(rst2!tot_wt_pricu_wt) * 0.4))
'        End If
'    End If
'    Set rst2 = Nothing
'
'End If
'
'If InStr(1, UCase(Equip), "BAR") > 0 Then
'    we = RateCal(Bartype, 1)
'End If
'
''++++++++++++++++++++
''5 Resin
''Dim epoxy As Double
'Set rst = New Recordset
'rst.Open "select distinct tot_wt_epoxy_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
'If rst.EOF = False And rst.BOF = False Then
'     we = prcCalResin(CDbl(rst.Fields(0)))
'End If
'Set rst = Nothing
'
'
'
'If InStr(1, UCase(Equip), "WOUND") > 0 Then
'
'    '6**************
'    'Bracket
'    we = RateCal(Brackett, 1)
'
'    '7**************
'    'Plate
'    we = RateCal(Plate, 1)
'
'Else
'
'    '6**************
'    'for insert
'    we = RateCal(CStr(WinBar), 4)
'
'    '7**************
'    'Plate
'    we = RateCal(Plate, 1)
'
'End If
'
''+++++++++++++++++++++++++++++++++
''Rest of the installation
'Set rst1 = New Recordset
'rst1.Open "costing", con, adOpenDynamic, adLockOptimistic
'rst1.AddNew
'rst1!Item_name = "-"
'rst1.Update
'
'Dim rstItem As Recordset
'
'Set rst = New Recordset
'rst.Open "select PRODUCT_CODE FROM MST_PART WHERE DRG_no = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'    P_CODE = rst.Fields(0)
'End If
'Set rst = Nothing
'
'
'
'
'Set rst = New Recordset
'rst.Open "select * from new_costmaster where PRODUCT_CODE = '" & P_CODE & "'", con, adOpenStatic, adLockReadOnly
''rst.Open "select * from costingcomm", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'    Do Until rst.EOF
'        rst1.AddNew
'        rst1!code = rst!item_code
'        Set rstItem = New Recordset
'        rstItem.Open "select * from MST_PART where PRODUCT_code = '" & rst!item_code & "'", con, adOpenStatic, adLockReadOnly
'            rst1!Item_name = rstItem!PRODUCT_DESCRIPTION & vbNullString
'            rst1!units = rstItem!uomC & vbNullString
'            If IsNull(rstItem!latest_rate) = True Then
'                rst1!Rate = 0
'            Else
'                rst1!Rate = Format(rstItem!latest_rate, "0.00")
'            End If
'            rst1!qty = Format(Trim(rst!qty), "0.00")
'            rst1!amount = Format(CDbl(rst!qty) * CDbl(rst1!Rate), "0.00")
'        Set rstItem = Nothing
'        rst1.Update
'        rst.MoveNext
'    Loop
'    'Sundary Material
'    rst1.AddNew
'    rst1!code = "----"
'    rst1!Item_name = "Sundry Material"
'    rst1!units = "Nos"
'    rst1!Rate = 100
'    rst1!qty = 1
'    rst1!amount = 100
'    rst1.Update
'    'Mould Charges
'    rst1.AddNew
'    rst1!code = "----"
'    rst1!Item_name = "Mould Charges"
'    rst1!units = "Nos"
'    rst1!Rate = 100
'    rst1!qty = 1
'    rst1!amount = 100
'    rst1.Update
'    'Packing Charges
'    rst1.AddNew
'    rst1!code = "----"
'    rst1!Item_name = "Packing Charges"
'    rst1!units = "Nos"
'    rst1!Rate = 100
'    rst1!qty = 1
'    rst1!amount = 100
'    rst1.Update
'
'End If
'Set rst = Nothing
'
'Set rst1 = Nothing
'
'err:
'    If Len(Trim(err.Description)) > 0 Then
'        MsgBox err.Description & vbCrLf & err.Number, vbCritical
'    End If
'
'End Function
'
'
'Private Function RateCal(item_code As String, CQty As Double)
'
'On Error GoTo err
'
''++++++++++++++++++++++++++++++++++++++
''Taking rate from costmaster
'
'Dim rs1 As Recordset
'
'Set rst = New Recordset
''rst.Open "select * from costmaster where part = '" & Item_code & "'", con, adOpenStatic, adLockReadOnly
'rst.Open "select * from mst_part where product_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'      Set rs1 = New Recordset
'      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'      rs1.AddNew
'      rs1!code = item_code
''      If InStr(rst!item_desc1 & vbNullString, "Insert") > 0 Then
''        If InStr(rst!item_desc1 & vbNullString, "M.S.") > 0 Then
''        'M.S.
''            If IsNull(rst!Type) = True Then
'                rs1!Item_name = rst!product_code & " " & rst!drg_no
''            Else
''                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!Mod & ")" & " Type:" & rst!Type
''            End If
''        Else
''        'Brass
''            If IsNull(rst!Type) = True Then
''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")"
''            Else
''                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''            End If
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "PLATE") > 0 Then
''      'Mounding Plate
''        If IsNull(rst!Type) = True Or Len(rst!Type) = 0 Then
''            If Len(Trim(rst!Mod)) = 0 Then
''                rs1!Item_name = "Mounting Plate " & rst!Drg
''            Else
''                rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")"
''            End If
''        Else
''            rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BRACKET") > 0 Then
''      'Mounting Bracket
''        If IsNull(rst!Type) = True And Len(Trim(rst!Type)) = 0 Then
''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")"
''        Else
''            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "TERMINAL") > 0 Then
''      'Mounting Bracket
''        If IsNull(rst!Type) = True Then
''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")"
''        Else
''            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!Mod & ")" & "Type:" & rst!Type
''        End If
''      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BUNCHED") > 0 Then
''      'Mounting Bracket
''        rs1!Item_name = rst!item_desc2 & " group copper bunch"
''      Else
''      'Else
''        rs1!Item_name = rst!item_desc1
''        If Len(Trim(rst!item_desc4)) > 0 Or IsNull(rst!item_desc4) = False Then
''            rs1!Item_name = rs1!Item_name & " " & rst!item_desc4
''        End If
''
''      End If
'
'      rs1!units = rst!uomC & vbNullString
'      rs1!Rate = Val(rst!latest_rate & vbNullString)
'      rs1!qty = CQty
'      rs1!amount = Val(rst!latest_rate & vbNullString) * CQty
'      rs1.Update
'      Set rs1 = Nothing
'
'End If
'Set rst = Nothing
'
'err:
'    If Len(Trim(err.Description)) > 0 Then
'        MsgBox err.Description & vbCrLf & err.Number, vbCritical
'    End If
'
'
'End Function
'
'
'Private Function prcCalResin(qty As Double)
'
'On Error GoTo err
'
''++++++++++++++++++++++++++++++++++++++
''For resin Only
'Set rst = New Recordset
''rst.Open "select * from costmaster where part = 'INS-00A'", con, adOpenStatic, adLockReadOnly
'rst.Open "select * from mst_part where product_code = '11010629-00'", con, adOpenStatic, adLockReadOnly
'
'If rst.EOF = False And rst.BOF = False Then
'
'     Set rs1 = New Recordset
'     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'
'     rs1.AddNew
'     rs1!code = "11010629-00"
'     rs1!Item_name = "Hot Setting Epoxy Resin Mix"
'     rs1!units = "Kgs"
'     rs1!Rate = rst!latest_rate
'     rs1!qty = CInt(qty)
'     rs1!amount = rst!latest_rate * CInt(qty)
'     rs1.Update
'     Set rs1 = Nothing
'
'End If
'Set rst = Nothing
'
'err:
'    If Len(Trim(err.Description)) > 0 Then
'        MsgBox err.Description & vbCrLf & err.Number, vbCritical
'    End If
'
'End Function
'
'Private Function prcCalMUCoreShape(TyprC As String, qty As Double)
'
'On Error GoTo err
'
''++++++++++++++++++++++++++++++++++++++
''For Core Only
'
'If TyprC = "R" Then
'    rrr = "CRGO Rectangular Maganetic Core"
'    rcode = "REC-00M"
'    RATEE = 2500
'Else
'    rrr = "CRGO Circular Maganetic Core"
'    rcode = "ROC-00M"
'    RATEE = 2500
'End If
'
'Dim rs1 As Recordset
'
'Set rst = New Recordset
'rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
'If rst.EOF = False And rst.BOF = False Then
'
'     Set rs1 = New Recordset
'     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'
'     rs1.AddNew
'     rs1!code = rcode
'     rs1!Item_name = rrr
'     rs1!units = "Kgs"
'     rs1!Rate = RATEE
'     rs1!qty = qty
'     rs1!amount = qty * RATEE
'     rs1.Update
'
'     Set rs1 = Nothing
'
'End If
'Set rst = Nothing
'
'err:
'    If Len(Trim(err.Description)) > 0 Then
'        MsgBox err.Description & vbCrLf & err.Number, vbCritical
'    End If
'
'End Function
'
'Private Function GetCorePrice(CoreShape As String, Qtykg As Double) As Currency
'
'    Dim rs As New ADODB.Recordset
'    Dim SQL1 As String
'    Set rs = New ADODB.Recordset
'    If CoreShape = "R" Then
'        If Mid(CRGO, 6, 6) = "A" Then
'            SQL1 = "CRE00A"
'        ElseIf Mid(CRGO, 6, 6) = "B" Then
'            SQL1 = "CRE00B"
'        ElseIf Mid(CRGO, 6, 6) = "C" Then
'            SQL1 = "CRE00C"
'        End If
'    Else
'        If Mid(CRGO, 6, 6) = "A" Then
'            SQL1 = "CRO00A"
'        ElseIf Mid(CRGO, 6, 6) = "B" Then
'            SQL1 = "CRO00B"
'        ElseIf Mid(CRGO, 6, 6) = "C" Then
'            SQL1 = "CRO00C"
'        End If
'    End If
'
'
'    rs.Open "Select * from Core_rate where coretype like '" & CoreShape & "%' and grade = '" & Mid(CRGO, 6, 6) & "'", con, adOpenKeyset, adLockReadOnly
''    rs.Open "select * from itemmast where item_code = '" & SQL1 & "'", con, adOpenStatic, adLockReadOnly
'    If rs.EOF = False And rs.BOF = False Then
'        GetCorePrice = rs!landedprice
'    End If
'    Set rs = Nothing
'End Function
'
'Private Function prcCalCoreShape(TyprC As String, qty As Double, CoreCode As String)
''++++++++++++++++++++++++++++++++++++++
''For Core Only
'
' If TyprC = "R" Then
'    rrr = CoreSz & " " & CRGO
'    rcode = CoreCode
'    RATEE = GetCorePrice("R", qty)
'Else
'    rrr = CoreSz & " " & CRGO
'    rcode = CoreCode
'    RATEE = GetCorePrice("C", qty)
'End If
'
'Dim rs1 As Recordset
'
'Dim rst1 As ADODB.Recordset
'Set rst1 = New Recordset
''rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
'rst1.Open "select * from MST_PART where Product_code = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
'
'If rst1.EOF = False And rst1.BOF = False Then
'     Set rs1 = New Recordset
'     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'
'     rs1.AddNew
'     rs1!code = rcode
'     rs1!Item_name = rrr
'     rs1!units = "Kgs"
'     rs1!Rate = rst1!latest_rate
'     rs1!qty = qty
'     rs1!amount = qty * rst1!latest_rate
'     rs1.Update
'
'     Set rs1 = Nothing
'Else
'     Set rs1 = New Recordset
'     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'
'     rs1.AddNew
'     rs1!code = "-"
'     rs1!Item_name = rrr
'     rs1!units = "Kgs"
'     rs1!Rate = 0
'     rs1!qty = qty
'     rs1!amount = qty * 0
'     rs1.Update
'
'     Set rs1 = Nothing
'
'End If
'
'Set rst1 = Nothing
'
'End Function
'
'Private Function RateCopper(item_code As String, CQty As Double)
''++++++++++++++++++++++++++++++++++++++
''Taking rate from costmaster
'
'Dim rs1 As Recordset
'
''item_code = "SWG " & item_code
'Set rst = New Recordset
''rst.Open "select * from costmaster where title like '" & Item_code & "%'", con, adOpenStatic, adLockReadOnly
'rst.Open "select * from MST_PART where material_type = 'SUPER ENAMELED ROUND COPPER WINDING WIRE;' and other_det_type = 'MEDIUM COVERING ;' and Dimension like '" & item_code & "%'", con, adOpenStatic, adLockReadOnly
'
''str1 = "select * from costmaster where title like '" & Item_code & "%'"
'If rst.EOF = False And rst.BOF = False Then
'
'      Set rs1 = New Recordset
'      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
'      rs1.AddNew
'      rs1!code = rst!product_code & vbNullString 'rst!part
'      rs1!Item_name = item_code & " Super Enamelled copper wire"
'      rs1!units = rst!uomC & vbNullString
'      rs1!Rate = rst!latest_rate
'      rs1!qty = CQty
'      rs1!amount = rst!latest_rate * CQty
'      rs1.Update
'      Set rs1 = Nothing
'
'End If
'Set rst = Nothing
'
'End Function
'
'
'
Dim CRGO As String
Dim CoreSz As String
Dim rst As Recordset

Public Function FindCoreID(DesCode As String, CShape As String, CorNo As Integer) As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim DWL, WWW, DOL, DOW, DHT As String
    Dim DID, DOD As String
    If CShape = "R" Then
    'For Rectangular
        
        Set rs = New ADODB.Recordset
        rs.Open "select ol,ow,wl,ww,ht from designspec where des_code = '" & DesCode & "' and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
        If rs.EOF = False And rs.BOF = False Then
'            If rs!WW > rs!WL Then
                DWw = CStr(rs!WW)
                DWL = CStr(rs!WL)
                
                DOL = CStr(rs!OL)
                DOW = CStr(rs!OW)
                
                DHT = CStr(rs!HT)
'            Else
'                DWL = CStr(rs!WW)
'                DWW = CStr(rs!WL)
'
'                DOL = CStr(rs!OW)
'                DOW = CStr(rs!OL)
'
'                DHT = CStr(rs!HT)
'            End If
            Dim y As Double
            Dim X As Integer
            If Len(DWL) = 4 Then
                Size = DWL
            ElseIf Len(DWL) = 3 Then
                Size = DWL & "0"
            ElseIf Len(DWL) = 2 Then
                Size = "0" & DWL & "0"
            ElseIf Len(DWL) = 1 Then
                Size = "00" & DWL & "0"
            End If
            
            If Len(DWw) = 4 Then
                Size = Size & DWw
            ElseIf Len(DWw) = 3 Then
                Size = Size & DWw & "0"
            ElseIf Len(DWw) = 2 Then
                Size = Size & "0" & DWw & "0"
            ElseIf Len(DWw) = 1 Then
                Size = Size & "00" & DWw & "0"
            End If
            
            
            y = (CDbl(DOL) - CDbl(DWL)) / 2
            'txtOOd.Text = CDbl(DWW) + y * 2
        '    MsgBox CInt(X)
            If InStr(1, CStr(y), ".") > 0 Then
                    X = y * 10
                Else
                    X = y
            End If
            If Len(CStr(X)) = 4 Then
                Size = Size & X
            ElseIf Len(CStr(X)) = 3 Then
                Size = Size & "0" & X
            ElseIf Len(CStr(X)) = 2 Then
                Size = Size & "0" & X & "0"
            ElseIf Len(CStr(X)) = 1 Then
                Size = Size & "00" & X & "0"
            End If
            
            If Len(DHT) = 4 Then
                Size = Size & DHT
            ElseIf Len(DHT) = 3 Then
                Size = Size & DHT & "0"
            ElseIf Len(DHT) = 2 Then
                Size = Size & "0" & DHT & "0"
            ElseIf Len(DHT) = 1 Then
                Size = Size & "00" & DHT & "0"
            End If
            
'            MsgBox Size
            Dim rst As New ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
            CoreSz = DWL & "X" & DWw & " " & DOL & "X" & DOW & " " & DHT
            If rst.EOF = False And rst.BOF = False Then
'                MsgBox "The item you are searching for is : " & "  " & rst!Item_code
                FindCoreID = rst!item_code
            Else
                '"""""""
                'reversing the size and checking again
                DWL = CStr(rs!WW)
                DWw = CStr(rs!WL)

                DOL = CStr(rs!OW)
                DOW = CStr(rs!OL)

                DHT = CStr(rs!HT)
                If Len(DWL) = 4 Then
                    Size = DWL
                ElseIf Len(DWL) = 3 Then
                    Size = DWL & "0"
                ElseIf Len(DWL) = 2 Then
                    Size = "0" & DWL & "0"
                ElseIf Len(DWL) = 1 Then
                    Size = "00" & DWL & "0"
                End If
                
                If Len(DWw) = 4 Then
                    Size = Size & DWw
                ElseIf Len(DWw) = 3 Then
                    Size = Size & DWw & "0"
                ElseIf Len(DWw) = 2 Then
                    Size = Size & "0" & DWw & "0"
                ElseIf Len(DWw) = 1 Then
                    Size = Size & "00" & DWw & "0"
                End If
                
                
                y = (CDbl(DOL) - CDbl(DWL)) / 2
                'txtOOd.Text = CDbl(DWW) + y * 2
            '    MsgBox CInt(X)
                If InStr(1, CStr(y), ".") > 0 Then
                        X = y * 10
                    Else
                        X = y
                End If
                If Len(CStr(X)) = 4 Then
                    Size = Size & X
                ElseIf Len(CStr(X)) = 3 Then
                    Size = Size & "0" & X
                ElseIf Len(CStr(X)) = 2 Then
                    Size = Size & "0" & X & "0"
                ElseIf Len(CStr(X)) = 1 Then
                    Size = Size & "00" & X & "0"
                End If
                
                If Len(DHT) = 4 Then
                    Size = Size & DHT
                ElseIf Len(DHT) = 3 Then
                    Size = Size & DHT & "0"
                ElseIf Len(DHT) = 2 Then
                    Size = Size & "0" & DHT & "0"
                ElseIf Len(DHT) = 1 Then
                    Size = Size & "00" & DHT & "0"
                End If
                
'                MsgBox Size
                'Dim rst As New ADODB.Recordset
                Set rst = New ADODB.Recordset
                rst.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
                CoreSz = DWw & "X" & DWL & " " & DOW & "X" & DOL & " " & DHT
                If rst.EOF = False And rst.BOF = False Then
'                    MsgBox "The item you are searching for is : " & "  " & rst!Item_code
                    FindCoreID = rst!item_code
                End If
                    
                '"""""""
            End If
            
        End If
        'If found the Code
        
        
        'Else generate new
    
    Else
    'For Circular
        Set rs = New ADODB.Recordset
        rs.Open "select OD,ID,ht from designspec where des_code = '" & DesCode & "'  and core_no = " & CorNo & "", con, adOpenKeyset, adLockReadOnly
        
        If rs.EOF = False And rs.BOF = False Then
'                If rs!ID > rs!od Then
                    DID = CStr(rs!oD)
                    DOD = CStr(rs!ID)
                    DHT = CStr(rs!HT)
'                Else
'                    DID = CStr(rs!ID)
'                    DOD = CStr(rs!od)
'                    DHT = CStr(rs!HT)
'                End If
                Size = ""
                
                If Len(DID) = 4 Then
                    Size = DID
                ElseIf Len(DID) = 3 Then
                    Size = DID & "0"
                ElseIf Len(DID) = 2 Then
                    Size = "0" & DID & "0"
                ElseIf Len(DID) = 1 Then
                    Size = "00" & DID & "0"
                End If
                
                If Len(DOD) = 4 Then
                    Size = Size & DOD
                ElseIf Len(DOD) = 3 Then
                    Size = Size & DOD & "0"
                ElseIf Len(DOD) = 2 Then
                    Size = Size & "0" & DOD & "0"
                ElseIf Len(DOD) = 1 Then
                    Size = Size & "00" & DOD & "0"
                End If
                
                
                If Len(DHT) = 3 Then
                    Size = Size & DHT & "0"
                ElseIf Len(DHT) = 2 Then
                    Size = Size & "0" & DHT & "0"
                ElseIf Len(DHT) = 1 Then
                    Size = Size & "000" & DHT
                End If
            
'            MsgBox Size
'            Dim rst As New ADODB.Recordset
            Dim rst1 As ADODB.Recordset
            Set rst1 = New ADODB.Recordset
            rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
            CoreSz = DID & " X " & DOD & "  " & DHT
            If rst1.EOF = False And rst1.BOF = False Then
'                MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
                FindCoreID = rst1!item_code
            Else
                'Reversing Size and checking again
                    DID = CStr(rs!ID)
                    DOD = CStr(rs!oD)
                    DHT = CStr(rs!HT)
                    Size = ""
                    
                    If Len(DID) = 4 Then
                        Size = DID
                    ElseIf Len(DID) = 3 Then
                        Size = DID & "0"
                    ElseIf Len(DID) = 2 Then
                        Size = "0" & DID & "0"
                    ElseIf Len(DID) = 1 Then
                        Size = "00" & DID & "0"
                    End If
                    
                    If Len(DOD) = 4 Then
                        Size = Size & DOD
                    ElseIf Len(DOD) = 3 Then
                        Size = Size & DOD & "0"
                    ElseIf Len(DOD) = 2 Then
                        Size = Size & "0" & DOD & "0"
                    ElseIf Len(DOD) = 1 Then
                        Size = Size & "00" & DOD & "0"
                    End If
                    
                    
                    If Len(DHT) = 3 Then
                        Size = Size & DHT & "0"
                    ElseIf Len(DHT) = 2 Then
                        Size = Size & "0" & DHT & "0"
                    ElseIf Len(DHT) = 1 Then
                        Size = Size & "000" & DHT
                    End If
                
'                MsgBox Size
    '            Dim rst As New ADODB.Recordset
                'Dim rst1 As ADODB.Recordset
                Set rst1 = New ADODB.Recordset
                rst1.Open "Select * from itemmast where item_desc4 = '" & Trim(Size) & "'", con, adOpenDynamic, adLockOptimistic
                CoreSz = DID & " X " & DOD & "  " & DHT
                If rst1.EOF = False And rst1.BOF = False Then
'                    MsgBox "The item you are searching for is : " & "  " & rst1!Item_code
                    FindCoreID = rst1!item_code
                End If
                'End Reversing Size and checking again
            End If
            Set rst1 = Nothing
            
            
        End If
    End If

End Function

Public Function Coating(DES As String)
Dim EqType As String

On Error GoTo err

'++++++++++++++++++++++++++++++++++++
'Deleteing coasting
Set rst = New Recordset
rst.Open "delete from costing", con, adOpenDynamic, adLockOptimistic
 Set rst = Nothing

Dim rst1 As Recordset

Dim priCross As String
Dim PTurns As Integer
Dim CopTer1 As String
Dim CopTer2 As String
Dim Bar1 As String
Dim Bar2 As String
Dim Brackett As String
Dim Plate As String
Dim Drg As String
Dim drgtype As String
Dim Equip As String
Dim CoreShape As String
Dim CoreMat As String
Dim CoreID As String
Dim PriCond As String
Dim PriCondType As String
Dim PriConductor As String

Dim Bartype As String

'+++++++++++++++++++++++++++++++++++++++++++++
'Fing Pric Cross Sction
Set rst = New Recordset
rst.Open "select coreshape, pri_cond_style, equip, pri_cross_sec, pri_turns , drg_no , drg_type, pri_cond from designspec where des_code = '" & DES & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    priCross = rst!pri_cross_sec
    PTurns = rst!pri_turns
    Drg = rst!drg_no & vbNullString
    drgtype = rst!DRG_TYPE & vbNullString
    Equip = rst!Equip & vbNullString
    CoreShape = rst!CoreShape & vbNullString
    PriConductor = rst!pri_cond & vbNullString
    If IsNull(rst!pri_cond_style) = False And Len(Trim(rst!pri_cond_style)) > 0 Then
        PriCond = Mid(rst!pri_cond_style, 1, 1)
        PriCondType = Mid(rst!pri_cond_style, InStr(1, rst!pri_cond_style, "-") + 1)
    End If
End If
Set rst = Nothing

If InStr(1, Equip, "Wound") > 0 Then
    EqType = "2"
ElseIf InStr(1, Equip, "Window") > 0 Then
    EqType = "1"
ElseIf InStr(1, Equip, "Bar") > 0 Then
    EqType = "1"
End If


Dim Single_Or_Double As String
If InStr(1, PriCondType, "Double") > 0 Then
    Single_Or_Double = "Double"
ElseIf InStr(1, PriCondType, "Single") > 0 Then
    Single_Or_Double = "Single"
ElseIf InStr(1, PriCondType, "Series") > 0 Then
    Single_Or_Double = "Series"
End If

If Single_Or_Double = "Double" Then
    If InStr(1, PriCondType, "Equal") > 0 Then
        Single_Or_Double = "Double-Equal"
    ElseIf InStr(1, PriCondType, "Different") > 0 Then
        Single_Or_Double = "Double-Different"
    Else
        Single_Or_Double = "Double"
    End If
End If

'+++++++++++++++++++++++++++++++++++++++++++++
'Finding code from primarycond11  -by providing c/s and turns
'++++++++++++++++++++++++++++++
'+changing ri primarycond from primartcond11
Set rst = New Recordset
str1 = "select * from primarycond11 where t_csp = '" & priCross & "' and t_np = '" & PTurns & "' and style = '" & Single_Or_Double & "'"
If Trim(priCross) = "" Then
    PriCond = 0
    rst.Open "select *  from primarycond11 where [t_np] = " & PTurns & " and [type] = " & PriCond & " and [style] = '" & PriCondType & "' and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and [style] = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
Else
'    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and type = '" & PriCond & "%' and style = '" & PriCondType & "%' and drg = '" & Drg & "' and mod = '" & DrgType & "'", con, adOpenStatic, adLockReadOnly
    strr = "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and mod = '" & drgtype & "' and style = '" & Single_Or_Double & "'"
    rst.Open "select *  from primarycond11 where t_csp = " & priCross & " and t_np = " & PTurns & " and drg = '" & Drg & "' and [mod] = '" & drgtype & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
End If
If rst.EOF = False And rst.BOF = False Then
    CopTer1 = rst!t_terp1 & vbNullString
    CopTer2 = rst!t_terp2 & vbNullString
    Bar1 = rst!cond_code1 & vbNullString
    If Len(rst!cond_code2) > 0 And IsNull(rst!cond_code2) = False Then
        Bar2 = rst!cond_code2
    End If
Else
    Set rst = New Recordset
    rst.Open "select *  from primarycond11 where t_np = " & PTurns & " and [drg] = '" & Drg & "' and [mod] = '" & drgtype & "' and t_cond_siz = '" & PriConductor & "' and style = '" & Single_Or_Double & "'", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        CopTer1 = rst!t_terp1 & vbNullString
        CopTer2 = rst!t_terp2 & vbNullString
        Bar1 = rst!cond_code1 & vbNullString
        If Len(rst!cond_code2) > 0 Or IsNull(rst!cond_code2) = False Then
            Bar2 = rst!cond_code2
        End If
    End If
End If
Set rst = Nothing


'+++++++++++++++++++++++++++++++++++++++++++++
'Finding code from MouldVol -by providing drg, type
Set rst = New Recordset
If Len(Trim(drgtype)) > 0 Then
    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "' and t_drgtype = '" & drgtype & "'", con, adOpenStatic, adLockReadOnly
Else
    rst.Open "select * from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
End If
If rst.EOF = False And rst.BOF = False Then
    If InStr(1, UCase(Equip), "WOUND") > 0 Then
        Brackett = rst!t_bktcode & vbNullString  'for wound
        Plate = rst!t_pltcode & vbNullString     'for wound
    Else
        WinBar = rst!t_mtginst & vbNullString   'for window and  Bar
        Plate = rst!t_pltcode & vbNullString
        If InStr(1, UCase(Equip), "BAR") > 0 Then
            Bartype = rst!t_buscode & vbNullString
        End If
    End If
End If
Set rst = Nothing

'ADD INF TO REPORT

Dim rss As Recordset

''Core material
'Set rst = New Recordset
'rst.Open "select distinct core_mat from designspec where des_code = '" & DES & "'", con, adOpenStatic
'If rst.EOF = False And rst.BOF = False Then
'    CoreMat = rst!core_mat
'End If
'Set rst = Nothing

Dim epoxy As Double
Dim rstCRGO As Recordset

Set rst = New Recordset
rst.Open "select distinct core_no, tot_wt_core_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
If rst.EOF = False And rst.BOF = False Then
'    If CoreMat = "MU metal" Then
'        '1***************
'        If CoreShape = "Rectangluar" Then
'            we = prcCalMUCoreShape("R", rst!tot_wt_core_wt)
'        Else
'            we = prcCalMUCoreShape("C", rst!tot_wt_core_wt)
'        End If
'    Else
        '1***************
                
        Do Until rst.EOF
            Set rstCRGO = New Recordset
            rstCRGO.Open "select core_mat, core_wt from designcore where dec_code = '" & DES & "' and coreno = " & rst!core_no & "", con, adOpenStatic, adLockReadOnly
            If rstCRGO.EOF = False And rstCRGO.BOF = False Then CRGO = rstCRGO.Fields(0) & vbNullString
            
            If CoreShape = "Rectangluar" Then
                CoreID = FindCoreID(DES, "R", rst!core_no)
                we = prcCalCoreShape("R", rstCRGO!core_wt, CoreID)
            Else
                CoreID = FindCoreID(DES, "C", rst!core_no)
                we = prcCalCoreShape("C", rstCRGO!core_wt, CoreID)
            End If
            Set rstCRGO = Nothing
            rst.MoveNext
        Loop
'    End If
End If
Set rst = Nothing

'2************************
'Super Enabled Copper
Set rst = New Recordset
rst.Open "select max(coreno) from designcore where dec_code = '" & DES & "'", con, adOpenStatic
If rst.EOF = False And rst.BOF = False Then
    X = rst.Fields(0)
End If
Set rst = Nothing

    'Finding swg of each core fro design core
For i = 1 To X
    Set rst = New Recordset
    rst.Open "select swg, sec_cu_wt from designcore where dec_code = '" & DES & "' and coreno = " & i & " ", con, adOpenStatic
    If rst.EOF = False And rst.BOF = False Then
        y = RateCopper(rst!SWG, rst!sec_cu_wt)
    End If
    Set rst = Nothing
Next

'3***************
'Copper bunch 1 or 2
Set rst = New Recordset
rst.Open "select distinct coreno from designcore where dec_code = '" & DES & "'", con, adOpenStatic
If rst.EOF = False And rst.BOF = False Then
    vall = rst!coreno
End If
Set rst = Nothing

If InStr(1, UCase(Equip), "WOUND") > 0 Then
    '4**************
    'Terminal
    we = RateCal(CopTer1, 1)
    we = RateCal(CopTer2, 1)
    
    
    
    '' +++++++++++++
    ''Copper strip tot_wt_pricu_wt
    Dim rst2 As Recordset
    Set rst2 = New Recordset
    rst2.Open "select tot_wt_pricu_wt  from designspec where des_code = '" & DES & "'", con, adOpenStatic
    If rst2.EOF = False And rst2.BOF = False Then
        ''if only one group present all wt. taken by 1
        If Bar2 = "" Then
            we = RateCal(Bar1, CDbl(rst2!tot_wt_pricu_wt))
        ''else wt. distributed
        Else
            we = RateCal(Bar1, (CDbl(rst2!tot_wt_pricu_wt) * 0.6))
            we = RateCal(Bar2, (CDbl(rst2!tot_wt_pricu_wt) * 0.4))
        End If
    End If
    Set rst2 = Nothing
    
End If

If InStr(1, UCase(Equip), "BAR") > 0 Then
    we = RateCal(Bartype, 1)
End If

'++++++++++++++++++++
'5 Resin
'Dim epoxy As Double
Set rst = New Recordset
rst.Open "select distinct tot_wt_epoxy_wt from designspec where des_code = '" & DES & "'", con, adOpenStatic
If rst.EOF = False And rst.BOF = False Then
     we = prcCalResin(CDbl(rst.Fields(0)))
End If
Set rst = Nothing



If InStr(1, UCase(Equip), "WOUND") > 0 Then

    '6**************
    'Bracket
    we = RateCal(Brackett, 1)
    
    '7**************
    'Plate
    we = RateCal(Plate, 1)

Else

    '6**************
    'for insert
    we = RateCal(CStr(WinBar), 4)
    
    '7**************
    'Plate
    we = RateCal(Plate, 1)

End If

'+++++++++++++++++++++++++++++++++
'Rest of the installation
Set rst1 = New Recordset
rst1.Open "costing", con, adOpenDynamic, adLockOptimistic
rst1.AddNew
rst1!Item_name = "-"
rst1.Update

Dim rstItem As Recordset


Set rst = New Recordset
rst.Open "select * from new_costmaster where name = '" & EqType & "'", con, adOpenStatic, adLockReadOnly
'rst.Open "select * from costingcomm", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
    Do Until rst.EOF
        rst1.AddNew
        rst1!code = rst!item_code
        Set rstItem = New Recordset
        rstItem.Open "select * from itemmast where item_code = '" & rst!item_code & "'", con, adOpenStatic, adLockReadOnly
            rst1!Item_name = rstItem!item_desc1 & vbNullString
            rst1!units = rstItem!uom & vbNullString
            If IsNull(rstItem!landedprice) = True Then
                rst1!Rate = 0
            Else
                rst1!Rate = Format(rstItem!landedprice, "0.00")
            End If
            rst1!qty = Format(Trim(rst!qty), "0.00")
            rst1!amount = Format(CDbl(rst!qty) * CDbl(rst1!Rate), "0.00")
        Set rstItem = Nothing
        rst1.Update
        rst.MoveNext
    Loop
    'Sundary Material
    rst1.AddNew
    rst1!code = "----"
    rst1!Item_name = "Sundry Material"
    rst1!units = "Nos"
    rst1!Rate = 100
    rst1!qty = 1
    rst1!amount = 100
    rst1.Update
    'Mould Charges
    rst1.AddNew
    rst1!code = "----"
    rst1!Item_name = "Mould Charges"
    rst1!units = "Nos"
    rst1!Rate = 100
    rst1!qty = 1
    rst1!amount = 100
    rst1.Update
    'Packing Charges
    rst1.AddNew
    rst1!code = "----"
    rst1!Item_name = "Packing Charges"
    rst1!units = "Nos"
    rst1!Rate = 100
    rst1!qty = 1
    rst1!amount = 100
    rst1.Update
    
End If
Set rst = Nothing

Set rst1 = Nothing

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Function


Private Function RateCal(item_code As String, CQty As Double)

On Error GoTo err

'++++++++++++++++++++++++++++++++++++++
'Taking rate from costmaster

Dim rs1 As Recordset

Set rst = New Recordset
'rst.Open "select * from costmaster where part = '" & Item_code & "'", con, adOpenStatic, adLockReadOnly
rst.Open "select * from itemmast where item_code = '" & item_code & "'", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then
      
      Set rs1 = New Recordset
      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
      rs1.AddNew
      rs1!code = item_code
      If InStr(rst!item_desc1 & vbNullString, "Insert") > 0 Then
        If InStr(rst!item_desc1 & vbNullString, "M.S.") > 0 Then
        'M.S.
            If IsNull(rst!Type) = True Then
                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!mod & ")"
            Else
                rs1!Item_name = "M.S. Insert " & rst!Drg & "(" & rst!mod & ")" & " Type:" & rst!Type
            End If
        Else
        'Brass
            If IsNull(rst!Type) = True Then
                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!mod & ")"
            Else
                rs1!Item_name = "Brass Insert " & rst!Drg & "(" & rst!mod & ")" & "Type:" & rst!Type
            End If
        End If
      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "PLATE") > 0 Then
      'Mounding Plate
        If IsNull(rst!Type) = True Or Len(rst!Type) = 0 Then
            If Len(Trim(rst!mod)) = 0 Then
                rs1!Item_name = "Mounting Plate " & rst!Drg
            Else
                rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!mod & ")"
            End If
        Else
            rs1!Item_name = "Mounting Plate " & rst!Drg & "(" & rst!mod & ")" & "Type:" & rst!Type
        End If
      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BRACKET") > 0 Then
      'Mounting Bracket
        If IsNull(rst!Type) = True And Len(Trim(rst!Type)) = 0 Then
            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!mod & ")"
        Else
            rs1!Item_name = "Mounting Bracket " & rst!Drg & "(" & rst!mod & ")" & "Type:" & rst!Type
        End If
      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "TERMINAL") > 0 Then
      'Mounting Bracket
        If IsNull(rst!Type) = True Then
            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!mod & ")"
        Else
            rs1!Item_name = "Copper Terminal " & rst!Drg & "(" & rst!mod & ")" & "Type:" & rst!Type
        End If
      ElseIf InStr(UCase(rst!item_desc1 & vbNullString), "BUNCHED") > 0 Then
      'Mounting Bracket
        rs1!Item_name = rst!item_desc2 & " group copper bunch"
      Else
      'Else
        rs1!Item_name = rst!item_desc1
        If Len(Trim(rst!item_desc4)) > 0 Or IsNull(rst!item_desc4) = False Then
            rs1!Item_name = rs1!Item_name & " " & rst!item_desc4
        End If
       
      End If
      
      rs1!units = rst!uom & vbNullString
      rs1!Rate = Val(rst!landedprice & vbNullString)
      rs1!qty = CQty
      rs1!amount = Val(rst!landedprice & vbNullString) * CQty
      rs1.Update
      Set rs1 = Nothing
            
End If
Set rst = Nothing

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If


End Function


Private Function prcCalResin(qty As Double)

On Error GoTo err

'++++++++++++++++++++++++++++++++++++++
'For resin Only
Set rst = New Recordset
'rst.Open "select * from costmaster where part = 'INS-00A'", con, adOpenStatic, adLockReadOnly
rst.Open "select * from itemmast where item_code = 'INS00A'", con, adOpenStatic, adLockReadOnly

If rst.EOF = False And rst.BOF = False Then
            
     Set rs1 = New Recordset
     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
     
     rs1.AddNew
     rs1!code = "INS00A"
     rs1!Item_name = "Hot Setting Epoxy Resin Max"
     rs1!units = "Kgs"
     rs1!Rate = rst!landedprice
     rs1!qty = CInt(qty)
     rs1!amount = rst!landedprice * CInt(qty)
     rs1.Update
     Set rs1 = Nothing
            
End If
Set rst = Nothing

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Function

Private Function prcCalMUCoreShape(TyprC As String, qty As Double)

On Error GoTo err

'++++++++++++++++++++++++++++++++++++++
'For Core Only

If TyprC = "R" Then
    rrr = "CRGO Rectangular Maganetic Core"
    rcode = "REC-00M"
    RATEE = 2500
Else
    rrr = "CRGO Circular Maganetic Core"
    rcode = "ROC-00M"
    RATEE = 2500
End If

Dim rs1 As Recordset

Set rst = New Recordset
rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
If rst.EOF = False And rst.BOF = False Then
            
     Set rs1 = New Recordset
     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
     
     rs1.AddNew
     rs1!code = rcode
     rs1!Item_name = rrr
     rs1!units = "Kgs"
     rs1!Rate = RATEE
     rs1!qty = qty
     rs1!amount = qty * RATEE
     rs1.Update
     
     Set rs1 = Nothing
            
End If
Set rst = Nothing

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Function

Private Function GetCorePrice(CoreShape As String, Qtykg As Double) As Currency
            
    Dim rs As New ADODB.Recordset
    Dim SQL1 As String
    Set rs = New ADODB.Recordset
    If CoreShape = "R" Then
        If Mid(CRGO, 6, 6) = "A" Then
            SQL1 = "CRE00A"
        ElseIf Mid(CRGO, 6, 6) = "B" Then
            SQL1 = "CRE00B"
        ElseIf Mid(CRGO, 6, 6) = "C" Then
            SQL1 = "CRE00C"
        End If
    Else
        If Mid(CRGO, 6, 6) = "A" Then
            SQL1 = "CRO00A"
        ElseIf Mid(CRGO, 6, 6) = "B" Then
            SQL1 = "CRO00B"
        ElseIf Mid(CRGO, 6, 6) = "C" Then
            SQL1 = "CRO00C"
        End If
    End If
    
    
    rs.Open "Select * from Core_rate where coretype like '" & CoreShape & "%' and grade = '" & Mid(CRGO, 6, 6) & "'", con, adOpenKeyset, adLockReadOnly
'    rs.Open "select * from itemmast where item_code = '" & SQL1 & "'", con, adOpenStatic, adLockReadOnly
    If rs.EOF = False And rs.BOF = False Then
        GetCorePrice = rs!landedprice
    End If
    Set rs = Nothing
End Function

Private Function prcCalCoreShape(TyprC As String, qty As Double, CoreCode As String)
'++++++++++++++++++++++++++++++++++++++
'For Core Only

 If TyprC = "R" Then
    rrr = CoreSz & " " & CRGO
    rcode = CoreCode
    RATEE = GetCorePrice("R", qty)
Else
    rrr = CoreSz & " " & CRGO
    rcode = CoreCode
    RATEE = GetCorePrice("C", qty)
End If

Dim rs1 As Recordset

Dim rst1 As ADODB.Recordset
Set rst1 = New Recordset
'rst.Open "select * from costmaster where part = '" & rcode & "'", con, adOpenStatic, adLockReadOnly
rst1.Open "select * from ITEMMAST where item_code = '" & rcode & "'", con, adOpenStatic, adLockReadOnly

If rst1.EOF = False And rst1.BOF = False Then
     Set rs1 = New Recordset
     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
     
     rs1.AddNew
     rs1!code = rcode
     rs1!Item_name = rrr
     rs1!units = "Kgs"
     rs1!Rate = RATEE
     rs1!qty = qty
     rs1!amount = qty * RATEE
     rs1.Update
     
     Set rs1 = Nothing
Else
     Set rs1 = New Recordset
     rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
     
     rs1.AddNew
     rs1!code = "-"
     rs1!Item_name = rrr
     rs1!units = "Kgs"
     rs1!Rate = RATEE
     rs1!qty = qty
     rs1!amount = qty * RATEE
     rs1.Update
     
     Set rs1 = Nothing

End If

Set rst1 = Nothing

End Function

Private Function RateCopper(item_code As String, CQty As Double)
'++++++++++++++++++++++++++++++++++++++
'Taking rate from costmaster

Dim rs1 As Recordset

item_code = "SWG " & item_code
Set rst = New Recordset
'rst.Open "select * from costmaster where title like '" & Item_code & "%'", con, adOpenStatic, adLockReadOnly
rst.Open "select * from itemmast where item_desc3 like '%" & item_code & "%'", con, adOpenStatic, adLockReadOnly

'str1 = "select * from costmaster where title like '" & Item_code & "%'"
If rst.EOF = False And rst.BOF = False Then
      
      Set rs1 = New Recordset
      rs1.Open "select * from costing", con, adOpenDynamic, adLockOptimistic
      rs1.AddNew
      rs1!code = rst!item_code & vbNullString 'rst!part
      rs1!Item_name = item_code & " Super Enamelled copper wire"
      rs1!units = rst!uom & vbNullString
      rs1!Rate = rst!landedprice
      rs1!qty = CQty
      rs1!amount = rst!landedprice * CQty
      rs1.Update
      Set rs1 = Nothing
            
End If
Set rst = Nothing

End Function


