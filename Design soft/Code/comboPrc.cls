VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comboPrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rst As Recordset
Dim Layers  As Double

Dim Odn, MLT, WHt, PCurrent, Toleng As Double

Public Function CalProtectingHt(BUP As Double, EArea As Double, CoreType As String, Mean As Double, SecCurr As Double, PCurrent As Double, STurn As Double, Odn As Double, ODs As Double, HT As Integer, WW As Integer, VA As Double) As Double
Dim BM As Double 'txtvk1
'txtDelta = FindDelta(Left(STC, InStr(1, UCase(STC), "K") - 1), txtSTurns(index-1), SecCurr, CDbl(txtODn(index-1)))
If Len(Odn) > 0 Then
    If UCase(CTfrm.txtNo) = "S" Then
        Odn = Odn * 1
'        Exit Sub
    ElseIf UCase(CTfrm.txtNo) = "D" Then
        Odn = Odn * 2
    ElseIf UCase(CTfrm.txtNo) = "T" Then
        Odn = Odn * 3
    ElseIf UCase(CTfrm.txtNo) = "Q" Then
        Odn = Odn * 4
    End If
End If
'DiaType = "D"
If InStr(1, CTfrm.txtCoreRatio, "-") = 0 Then
'For single primary
    If CTfrm.cmbStyle = "Rectangluar" Then
        Layers = CalLayers(ODs, 0, STurn, ODs, WW)
    Else
        Layers = CalLayersRound(ODs, 0, STurn, ODs, WW)
    End If
Else
'For multi pri
    If Len(ODs) > 0 Then
'        If DiaType = "D" Then
'            'For Different Dia
'            If cmbStyle = "Rectangluar" Then
'                Layers = CalLayersDiF(ODs, Index)
'            Else
'                Layers = CalLayersRoundDiF(ODs, Index)
'            End If
'        ElseIf DiaType = "S" Or DiaType = "" Then
            'For Same Dia
            If cmbStyle = "Rectangluar" Then
                Layers = CalLayers(ODs, 0, STurn, ODs, WW)
            Else
                Layers = CalLayersRound(ODs, 0, STurn, ODs, WW)
            End If
'        End If
    End If
End If
    Dim Lay As Integer
    
'If point found in the layers then + 1 layers
    If InStr(1, Layers, ".") > 0 Then
        Lay = Left(Layers, InStr(1, Layers, ".") - 1) + 1
    Else
        If Len(Layers) > 0 Then
            Lay = Layers
        Else
            Exit Function
        End If
    End If
    
    If DiaType = "D" Then
'Total Winding Ht for different Conductor
        If Index = 0 Then
            'GCross.TextMatrix(shibu, 8) = 1.5 + (GCross.TextMatrix(shibu, 1) * GCross.TextMatrix(shibu, 6)) + ((GCross.TextMatrix(shibu, 6) - 1) * 0.25) + (1.5 + 3)
            WHt = 1.5 + (ODs * Lay) + ((Lay - 1) * 0.25) + (1.5 + 3)
            WHt = Round(WHt, 3)
        Else
            'GCross.TextMatrix(shibu, 8) = (GCross.TextMatrix(shibu, 1) * (Count + 1)) + (((Count + 1) - 1) * 0.25) + (1.5 + 3)
            WHt = (ODs * Lay) + ((Lay - 1) * 0.25) + (1.5 + 3)
            WHt = Round(WHt, 3)
        End If
'Mean Length of Turns
        If Index = 0 Then
            MLT = (WHt * 3.142) + ((2 * BUP) + (2 * HT))
        Else
            MLT = PrcGcross(1, 2)
        End If
        
    Else
'Total Winding HT
        WHt = 1.5 + (ODs * Lay) + ((Lay - 1) * 0.25) + (1.5 + 3)
        WHt = Round(WHt, 3)
        
'Mean Length of Turns
    MLT = (WHt * 3.142) + ((2 * BUP) + (2 * HT))
    MLT = Round(MLT, 2)
        
    End If
    


    Dim mincurr As Double
    Dim value As String * 1
    mincurr = PCurrent
'    For i = 1 To Len(PCurrent) + 1
'    value = Mid(PCurrent, i, i)
'        If value = "-" Or i > Len(PCurrent) Then
'             If CDbl(str1) < mincurr Or mincurr = 0 Then
'                  mincurr = CDbl(str1)
'             End If
'             str1 = ""
'        Else
'            str1 = str1 & value
'        End If
'    Next

    
    If CTfrm.cmbStyle = "Rectangluar" Then
        mincurr = mincurr * CTfrm.txtturns / SecCurr
'        lowe = prcCalLower(CDbl(MinCurr), Index)
    Else
        mincurr = mincurr * CTfrm.txtturns / SecCurr
'        lowe = prcCalLowerCir(CDbl(MinCurr), Index)
    End If
'Total Length of wire
If DiaType = "S" Then

    Toleng = CDbl(MLT) * mincurr
Else
'LM * btween turns
    Toleng = CDbl(MLT) * ToTurns
End If
'Secondary Resistances
    '@ 20 Degee Cen
        Res = 0.0177 * (Toleng / 1000) / Odn
        Res = Round(Res, 3)
'@ 75 Degee Cen
        Rec75 = Res * 1.2161
        Rec75 = Round(Rec75, 3)
        
''Copper Weigth
'        txtCWt(Index) = ((CDbl(MLT) * STurns * Odn * 8.9) / 1000000) * 1.1
'        txtCWt(Index) = Round(txtCWt(Index), 2)
'MsgBox ToLeng
'If CoreType <> "Special Protection" Then


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'**********************FOR SAME CONDUCTOR
'==========================================
'    If DiaType = "S" Then
    
    '@ 75 Degee Cen
        Rec75 = Res * 1.2161
        Rec75 = Round(Rec75, 3)
        
        If CTfrm.cmbFreq = "50" Then
    'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
    
            BM50 = (VA + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / ((mincurr * SecCurr) * EArea)
            
            BM50 = Round(BM50, 0)
'            BM50.Visible = True
'            BM60.Text = ""
'            BM60.Visible = False
'            BM = Round(BM50, 0)
'            'Quatar VA Class 0.1, 0.2, 0.5  , 1
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA Class 3 ,5
'            BMatHalf = ((VA / 2) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM50
'                ISF = Round(ISF, 2)
'            End If
            CalProtectingHt = BM50
        ElseIf CTfrm.cmbFreq = "60" Then
    'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
            BM60 = (VA + ((SecCurr * SecCurr) * Rec75)) * 3.75 * 100000 / (mincurr * SecCurr * EArea)
            BM60 = Round(BM60, 0)
            CalProtectingHt = BM60
'            BM60.Visible = True
'            BM50.Text = ""
'            BM50.Visible = False
'            BM = BM60
'            'Quatar VA Class 0.1, 0.2, 0.5  , 1
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA Class 3 ,5
'            BMatHalf = ((VA / 2) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM60
'                ISF = Round(ISF, 2)
'            End If
        End If
        
        
        
        'CalLayers =
'        If cmbFreq = "50" Then
'    'FREQ @ 50 c1 = Aeff * no of turns / 3000
'            C1 = EArea * STurns / 3000
'            C1 = Round(C1, 3)
'        ElseIf cmbFreq = "60" Then
'    'FREQ @ 60c1 = Aeff * no of turns / 2500
'            C1 = EArea * STurns / 2500
'            C1 = Round(C1, 3)
'        End If
'
'        txtC2(Index) = CDbl(Mean) / CDbl(STurns)
    '    MsgBox txtC2(Index)
'        If InStr(1, txtC2(Index), "E") > 0 Then
'            txtC2(Index) = Round(txtC2(Index), 3)
'        End If
        'C2 with factor 1.3 * c2
'        txtC2(Index) = txtC2(Index) * 1.3
'        txtC2(Index) = Round(txtC2(Index), 3)
'
'        txtVkCal(Index) = C1 * 96
'        txtECurr(Index) = ""
'        C1.Text = ""
'        txtC2(Index).Text = ""
'    Else
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''**********************Different CONDUCTOR
''==========================================
'        Dim TotalRec1 As Double
'
'        '@ 75 Degee Cen
'        Rec75 = Res * 1.2161
'        Rec75 = Round(Rec75, 3)
'
'        TotalRec1 = 0
'        For l = 0 To Index
'            TotalRec1 = TotalRec1 + txtRec75(l)
'        Next
'        Rec75 = TotalRec1
'        If cmbFreq = "50" Then
'    'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
'
'            BM50 = (VA + ((SecCurr * SecCurr) * TotalRec1)) * 4.5 * 100000 / ((mincurr * SecCurr) * EArea)
'
'            BM50 = Round(BM50, 0)
'            BM50.Visible = True
'            BM60.Text = ""
'            BM60.Visible = False
'            BM = Round(BM50, 0)
'            'Quatar VA Class 0.1, 0.2, 0.5  , 1
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * TotalRec1)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA Class 3 ,5
'            BMatHalf = ((VA / 2) + ((SecCurr * SecCurr) * TotalRec1)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM50
'                ISF = Round(ISF, 2)
'            End If
'        ElseIf cmbFreq = "60" Then
'    'BM @ 60 = 3.75 * (VA + I*I*R) * 100000 / ATs * aeff
'            BM60 = (VA + (Sqr(SecCurr) * TotalRec1)) * 3.75 * 100000 / (mincurr * SecCurr * EArea)
'            BM60 = Round(BM60, 0)
'            BM60.Visible = True
'            BM50.Text = ""
'            BM50.Visible = False
'            BM = BM60
'            'Quatar VA Class 0.1, 0.2, 0.5  , 1
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * TotalRec1)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA Class 3 ,5
'            BMatHalf = ((VA / 2) + ((SecCurr * SecCurr) * TotalRec1)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM60
'                ISF = Round(ISF, 2)
'            End If
'        End If
'
'        If cmbFreq = "50" Then
'    'FREQ @ 50 c1 = Aeff * no of turns / 3000
'            C1 = EArea * STurns / 3000
'            C1 = Round(C1, 3)
'        ElseIf cmbFreq = "60" Then
'    'FREQ @ 60c1 = Aeff * no of turns / 2500
'            C1 = EArea * STurns / 2500
'            C1 = Round(C1, 3)
'        End If
'
'        txtC2(Index) = CDbl(Mean) / CDbl(STurns)
'    '    MsgBox txtC2(Index)
'        If InStr(1, txtC2(Index), "E") > 0 Then
'            txtC2(Index) = Round(txtC2(Index), 3)
'        End If
'        'C2 with factor 1.3 * c2
'        txtC2(Index) = txtC2(Index) * 1.3
'        txtC2(Index) = Round(txtC2(Index), 3)
'
'        txtVkCal(Index) = C1 * 96
'        txtECurr(Index) = ""
'        C1.Text = ""
'        txtC2(Index).Text = ""
'    End If
'Else
'    If DiaType = "S" Then
'        If cmbFreq = "50" Then
'    'FREQ @ 50 c1 = Aeff * no of turns / 3000
'            C1 = EArea * STurns / 3000
'            C1 = Round(C1, 3)
'        ElseIf cmbFreq = "60" Then
'    'FREQ @ 60c1 = Aeff * no of turns / 2500
'            C1 = EArea * STurns / 2500
'            C1 = Round(C1, 3)
'        End If
'
'        txtC2(Index) = CDbl(Mean) / CDbl(STurns)
'        If InStr(1, txtC2(Index), "E") > 0 Then
'            txtC2(Index) = Round(txtC2(Index), 3)
'        End If
'        'C2 with factor 1.3 * c2
'        txtC2(Index) = txtC2(Index) * 1.3
'        txtC2(Index) = Round(txtC2(Index), 3)
'
'        txtVkCal(Index) = C1 * 96
'        txtVkCal(Index) = Round(txtVkCal(Index), 0)
'
'        'Excitation Current various point
'        If txtV1 = "Vk/4" Then
'            txtECurr(Index) = 0.092 * txtC2(Index) * 1000
'        ElseIf txtV1 = "Vk/2" Then
'            txtECurr(Index) = 0.145 * txtC2(Index) * 1000
'        ElseIf txtV1 = "" Then
'            txtECurr(Index) = 0.34 * txtC2(Index) * 1000
'        End If
'    Else
'        For l = 0 To Index
'            TotalRec1 = TotalRec1 + txtRec75(l)
'        Next
'        Rec75 = TotalRec1
'        If cmbFreq = "50" Then
'    'FREQ @ 50 c1 = Aeff * no of turns / 3000
'            C1 = EArea * STurns / 3000
'            C1 = Round(C1, 3)
'        ElseIf cmbFreq = "60" Then
'    'FREQ @ 60c1 = Aeff * no of turns / 2500
'            C1 = EArea * STurns / 2500
'            C1 = Round(C1, 3)
'        End If
'
'        txtC2(Index) = CDbl(Mean) / CDbl(STurns)
'        If InStr(1, txtC2(Index), "E") > 0 Then
'            txtC2(Index) = Round(txtC2(Index), 3)
'        End If
'        'C2 with factor 1.3 * c2
'        txtC2(Index) = txtC2(Index) * 1.3
'        txtC2(Index) = Round(txtC2(Index), 3)
'
'        txtVkCal(Index) = C1 * 96
'        txtVkCal(Index) = Round(txtVkCal(Index), 0)
'
'        'Excitation Current various point
'        If txtV1 = "Vk/4" Then
'            txtECurr(Index) = 0.092 * txtC2(Index) * 1000
'        ElseIf txtV1 = "Vk/2" Then
'            txtECurr(Index) = 0.145 * txtC2(Index) * 1000
'        ElseIf txtV1 = "" Then
'            txtECurr(Index) = 0.34 * txtC2(Index) * 1000
'        End If
'
'    End If
'End If
''####For Accuracy
'If InStr(1, PCurrent, "-") > 0 Then
''Secondary Resistances
'    '@ 20 Degee Cen
'    Res = 0.0177 * (lowe / 1000) / Odn
'    Res = Round(Res, 3)
'
''@ 75 Degee Cen
'    Rec75 = Res * 1.2161
'    Rec75 = Round(Rec75, 3)
'
'''''BM @ 50
''''    txtBM50(index-1) = (VA + (Sqr(SecCurr) * txtrec75(index-1))) * 4.5 * 100000 / (MinCurr * EArea)
'
'    If CoreType <> "Special Protection" Then
'        If cmbFreq = "50" Then
'    'BM @ 50= 4.5 * (VA + I*I*R) * 100000 / ATs * aeff
'            MsgBox Sqr(SecCurr)
'            BM50 = (VA + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            BM50 = Round(BM50)
'            BM = BM50
'            'Quatar VA Class 0.1, 0.2, 0.5  , 1
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA Class 3 ,5
'            BMatHalf = ((VA / 2) + ((SecCurr * SecCurr) * Rec75)) * 4.5 * 100000 / (mincurr * SecCurr * EArea)
'            'Calculate ISF
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM50
'            Else
'    '####Calculation of Composite Error at Normal
'            Dim A, b, C As Double
'            C = CompNormal(BM50, Mean, mincurr)
'            EcatN = CompNormalPartII(C, SecCurr)
'    '####Calculation of Composite Error at ALF
'            A = CompALF(BM50, Mean, mincurr, ISF)
'            C = (BM50 * ISF)
'            EcatAlf = CompALFPartII(CDbl(A), SecCurr, C, ISF)
'            End If
'        ElseIf cmbFreq = "60" Then
'    'BM @ 60 = 3.5 * (VA + I*I*R) * 100000 / ATs * aeff
'            BM60 = (VA + ((SecCurr * SecCurr) * Rec75)) * 3.75 * 100000 / (mincurr * SecCurr * EArea)
'            BM60 = Round(BM60)
'            BM = BM60
'            'Quatar VA
'            BMatQuatar = ((VA / 4) + ((SecCurr * SecCurr) * Rec75)) * 3.75 * 100000 / (mincurr * SecCurr * EArea)
'            'Half VA
'            BMatHalf = ((VA / 2) + (Sqr(SecCurr * SecCurr) * Rec75)) * 3.75 * 100000 / (mincurr * SecCurr * EArea)
'            'ISF
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM60
'            Else
'        '####Calculation of Composite Error at Normal
'                A = CompNormal(BM60, Mean, STurns)
'                EcatN = CompNormalPartII(CDbl(A), SecCurr)
'        '####Calculation of Composite Error at ALF
'                A = CompALF(BM60, Mean, STurns, ISF)
'                b = BM60 * ISF
'                EcatAlf = CompALFPartII(CDbl(A), SecCurr, CDbl(b), ISF)
'            End If
'        End If
'    Else
'         If cmbFreq = "50" Then
'    'FREQ @ 50 c1 = Aeff * no of turns / 3000
'            C1 = EArea * STurns / 3000
'        ElseIf cmbFreq = "60" Then
'    'FREQ @ 60c1 = Aeff * no of turns / 2500
'            C1 = EArea * STurns / 2500
'        End If
'    End If
'Else
'    If UCase(CoreType) <> UCase("Special Protection") Then
'        If cmbFreq = "60" Then
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM60
'            Else
'        '####Calculation of Composite Error at Normal
'                A = CompNormal(BM60, Mean, STurns)
'        'Ie at Normal Current
'                IeNor = A
'                EcatN = CompNormalPartII(CDbl(A), SecCurr)
'        '####Calculation of Composite Error at ALF
'                A = CompALF(BM60, Mean, STurns, ISF)
'         'Ie as ALF
'                txtIeALF(Index) = A
'                b = BM60 * ISF
'                EcatAlf = CompALFPartII(CDbl(A), SecCurr, CDbl(b), ISF)
'            End If
'        Else
'            If CoreType = "Metering" Then
'                ISF = 21000 / BM50
'                ISF = Round(ISF, 2)
'            Else
'        '####Calculation of Composite Error at Normal
'                A = CompNormal(BM50, Mean, STurns)
'        'Ie at Normal Current set a
'                IeNor = A
'                IeNor = Round(IeNor, 3)
'                EcatN = CompNormalPartII(CDbl(A), SecCurr)
'        '####Calculation of Composite Error at ALF
'                A = CompALF(BM50, Mean, STurns, ISF)
'        'Ie at ALF
'                txtIeALF(Index) = A
'                txtIeALF(Index) = Round(txtIeALF(Index), 3)
'                b = BM50 * ISF
'                EcatAlf = CompALFPartII(CDbl(A), SecCurr, CDbl(b), ISF)
'            End If
'        End If
'    End If
'End If
'
End Function

'Calculating layers RECTANGULAR
Private Function CalLayers(Dia As Double, ind As Integer, STurns As Double, ODs As Double, WW1 As Integer) As Double

    Dim value As Double

    Dim Turnss As Double
'    MsgBox txtSTurns(Index)
    Turnss = STurns

    Dim Count As Byte
    Dim Ins(20) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 3
    Dim K As Double

'Type of Conductor
    If UCase(CTfrm.txtNo) = "S" Then
        Dia = Dia * 1
    ElseIf UCase(CTfrm.txtNo) = "D" Then
        Dia = Dia * 2
    ElseIf UCase(CTfrm.txtNo) = "T" Then
        Dia = Dia * 3
    ElseIf UCase(CTfrm.txtNo) = "Q" Then
        Dia = Dia * 4
    End If
'Counting No of turns

    Dim typ As Double

    For i = 0 To CInt(STurns)
        K = 0
'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (ODs * Count)
        End If

        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If

        i = ((CDbl(CTfrm.txtWL1) - 2 * (K)) * 0.9) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If

        value = value + i

'        If value >= CInt(txtSecTurn) Then
'             typ = 0.25
'             CalLayers = (Count + typ)
'             Exit Function
'        End If

'For secound limb

        value = value + i
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.5
'             CalLayers = (Count + typ)
'             Exit Function
'        End If

'For third limb

        Dim s As Double

        s = (((CDbl(WW1) - 2 * (K)) * 0.9) / Dia)
        'Rounding of s
        If InStr(1, s, ".") > 0 Then
            s = Left(s, InStr(1, s, ".") - 1)
        End If
        value = value + s
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.75
'             CalLayers = (Count + typ)
'             Exit Function
'        End If

'For Fourth Limb

        value = value + s
        turns1 = Turnss
        Turnss = Turnss - value

        If Turnss <= 0 Then
'            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                CalLayers = (Count + typ)
                Exit Function
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                CalLayers = (Count + typ)
                Exit Function
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1

    Next
    
    
    
End Function

'CalCulate layers Reound
Private Function CalLayersRound(Dia As Double, Index As Integer, STurns As Double, ODs As Double, ID As Integer) As Double
        
    Dim value As Double
    
    Dim Turnss As Double
'    MsgBox STurns
    Turnss = STurns
        
    Dim Count As Byte
    Dim Ins(20) As Double
    Ins(0) = 1.5
    Ins(1) = 0.25
    Ins(2) = 0.25
    Ins(3) = 0.25
    Ins(4) = 0.25
    Ins(5) = 0.25
    Ins(6) = 0.25
    Ins(7) = 0.25
    Ins(8) = 0.25
    Ins(9) = 0.25
    Ins(10) = 0.25
    Ins(11) = 0.25
    Ins(12) = 0.25
    Ins(13) = 0.25
    Ins(14) = 0.25
    Ins(15) = 0.25
    Ins(16) = 0.25
    Ins(17) = 0.25
    Ins(18) = 0.25
    Ins(19) = 0.25
    Ins(20) = 3
    Dim K As Double
    
'Type of Conductor
    If UCase(CTfrm.txtNo) = "S" Then
        Dia = Dia * 1
    ElseIf UCase(CTfrm.txtNo) = "D" Then
        Dia = Dia * 2
    ElseIf UCase(CTfrm.txtNo) = "T" Then
        Dia = Dia * 3
    ElseIf UCase(CTfrm.txtNo) = "Q" Then
        Dia = Dia * 4
    End If
'Counting No of turns
    
    Dim typ As Double
    
    For i = 0 To CInt(STurns)
        K = 0
'For First limb
        For l = 0 To Count
            K = K + Ins(l)
        Next
        If Count > 0 Then
            K = K + (ODs * Count)
        End If
        
        'Rounding of k
        If InStr(1, K, ".") > 0 Then
            If Mid(K, InStr(1, K, ".") + 1, 1) <= 5 Or Mid(K, InStr(1, K, ".") + 1, 1) = 0 Then
                K = Mid(K, 1, InStr(1, K, ".") - 1) & "." & 5
            Else
                K = Mid(K, 1, InStr(1, K, ".") - 1) + 1
            End If
        End If
        
        i = ((CDbl(ID) - 2 * (K)) * 0.9 * 3.142) / Dia
        'Rounding of i
        If InStr(1, i, ".") > 0 Then
            i = Left(i, InStr(1, i, ".") - 1)
        End If
        
        value = value + i
        
'        If value >= CInt(txtSecTurn) Then
'             typ = 0.25
'             CalLayers = (Count + typ)
'             Exit Function
'        End If
        
'For secound limb

'        value = value + i

        turns1 = Turnss
        Turnss = Turnss - value
        
        If Turnss <= 0 Then
'            MsgBox ((turns1 / value) * 100)
            finalt = ((turns1 / value) * 100)
            If finalt <= 25 Then
                typ = 0.25
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 25 And finalt <= 50 Then
                typ = 0.5
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 50 And finalt <= 75 Then
                typ = 0.75
                CalLayersRound = (Count + typ)
                Exit Function
            ElseIf finalt > 75 And finalt <= 100 Then
                typ = 1
                CalLayersRound = (Count + typ)
                Exit Function
            End If
        End If
        value = 0
        typ = 0
        Count = Count + 1

    Next
    
    CalLayersRound = i
    
End Function


Private Function PrcGcross(shibu As Integer, shibu1 As Integer) As Double

Dim W1, W2 As Double
W1 = 0
For i = 0 To shibu1
    W1 = W1 + Individualfrm.txtWHt(i)
Next
W2 = Individualfrm.txtWHt(shibu)
PrcGcross = ((W2 / 2) + W1) * (2 * 3.142) + (2 * txtBUP) + (2 * txtHt)

End Function
