VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Weights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset


Public Function CalResinWt(Drg As String, drgtype As String, SecCUWt As Double, CoreWt As Double, ITemType As String, PriCuWt As Double) As Double

On Error GoTo err

'+++++++++++++++++++++++Finding Mould
Dim Mould As Double
Set rst = New ADODB.Recordset
If Len(drgtype) = 0 Then
    rst.Open "select t_mould_vo from mouldvol2 where t_drg = '" & Drg & "'", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        Mould = rst.Fields(0)
    Else
        MsgBox "No Mould Volume found for Drg : " & Drg
        Mould = 0
    End If
Else
    rst.Open "select t_mould_vo from mouldvol2 where t_drg = '" & Drg & "' and t_DrgType = '" & drgtype & "' ", con, adOpenStatic, adLockReadOnly
    If rst.EOF = False And rst.BOF = False Then
        Mould = rst.Fields(0)
    Else
        MsgBox "No Mould Volume found for Drg : " & Drg & " Type : " & drgtype, vbInformation
'        Casting.Rows = 1
        CTfrm.txtCoreWT = ""
        CTfrm.txtSecCuWt = ""
        CTfrm.txtpricuWt = ""
        CTfrm.txtResinWt = ""
        Mould = 0
        Exit Function
    End If
End If
Set rst = Nothing

'+++++++++++++++++++++++Window Type / Bar Type
'Copper Vol(Cubic cm) = (SecCUWt / Density) * 1000
'CoreVol(Cubic cm) = (CoreWt / Density) * 1000 (Desity will Differ for MU metal)
'Calculate resin Vol(Cubic cm) = MouldVol - (CoreVol + SecCopperVol)
'Calculate resin Wt(Kg) = (ResinVol X ResinDen) / 1000

Dim CuVol, CoreVol, ResinVol, ResinWt, PriCuVol As Double
Dim Dens, ResinDen As Double

Dens = 7.65
ResinDen = 1.8

If InStr(1, ITemType, "Window") > 0 Or InStr(1, ITemType, "Bar") > 0 Then

    'Sec Copper Vol
    CuVol = (SecCUWt / 8.9) * 1000    '--X
    'Core Vol
    CoreVol = (CoreWt / Dens) * 1000  '--y
    'Resin Vol
    ResinVol = Mould - (CoreVol + CuVol) 'Mould-(x+y)
    'Resin Wt
    ResinWt = (ResinVol * ResinDen) / 1000
    ResinWt = Round(ResinWt, 0)
    If InStr(1, ITemType, "Bar") > 0 Then
        CTfrm.Label82.Caption = "PRI. CU. BAR WEIGHT (kg) :"
        ' To be taken from the mould vol database
'        CTfrm.txtPrimaryCopper = 0
    Else
        CTfrm.Label82.Caption = "PRIMARY CU WEIGHT (kg) :"
    End If
    
ElseIf InStr(1, ITemType, "Wound") > 0 Then
'+++++++++++++++++++++++Wound
'Copper Vol(Cubic cm) = (SecCUWt / Density) * 1000
'CoreVol(Cubic cm) = (CoreWt / Density) * 1000 (Desity will Differ for MU metal)
'Calculate resin Vol(Cubic cm) = MouldVol - (CoreVol + SecCopperVol + PriCUVol)
'Calculate resin Wt(Kg) = (ResinVol X ResinDen) / 1000

    'Sec Copper Vol
    CuVol = (SecCUWt / 8.9) * 1000    '--X
    'Core Vol
    CoreVol = (CoreWt / Dens) * 1000  '--y
    'PriVol
    PriCuVol = (PriCuWt / 8.9) * 1000 '--z
    'Resin Vol
    ResinVol = Mould - (CoreVol + CuVol) 'Mould-(x+y+z)
    'Resin Wt
    ResinWt = (ResinVol * ResinDen) / 1000
    ResinWt = Round(ResinWt, 0)
End If

CalResinWt = Round(ResinWt, 3)

err:
    If Len(Trim(err.Description)) > 0 Then
        MsgBox err.Description & vbCrLf & err.Number, vbCritical
    End If

End Function


