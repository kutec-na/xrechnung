VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMapRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsMapRow - Spaltenzuordnung eines Positionsfeldes (BG-25)
' BT-Nr. | Bezeichnung | Spalte (Auswahl) | Zeile(n) in der Zelle | Standardwert
'==============================================================================
Option Explicit

Private Const X_ID As Single = 2
Private Const X_NAME As Single = 44
Private Const X_COL As Single = 192
Private Const X_LINES As Single = 306
Private Const X_DEF As Single = 342
Private Const W_DEF As Single = 54
Private Const CTL_H As Single = 15

Public Owner As Object
Public id As String
Public req As String
Public kind As String

Public lblID As MSForms.Label
Public lblName As MSForms.Label
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Public WithEvents txtLines As MSForms.TextBox
Attribute txtLines.VB_VarHelpID = -1
Public WithEvents txtDef As MSForms.TextBox
Attribute txtDef.VB_VarHelpID = -1

Private mBusy As Boolean

Public Sub Create(ByVal frm As Object, ByVal container As Object, ByVal y As Single, _
                  ByVal sID As String, ByVal sName As String, ByVal sReq As String, ByVal sKind As String)
    Set Owner = frm
    id = sID
    req = sReq
    kind = sKind

    Set lblID = container.Controls.Add("Forms.Label.1")
    lblID.Caption = sID
    lblID.Left = X_ID: lblID.Top = y + 2: lblID.Width = X_NAME - X_ID - 2: lblID.Height = CTL_H - 2
    If sReq = "P" Then lblID.Font.Bold = True

    Set lblName = container.Controls.Add("Forms.Label.1")
    lblName.Caption = sName
    lblName.WordWrap = False
    lblName.ControlTipText = sID & " - " & sName & IIf(sReq = "P", " (Pflicht)", "")
    lblName.Left = X_NAME: lblName.Top = y + 2: lblName.Width = X_COL - X_NAME - 4: lblName.Height = CTL_H - 2

    Set cbo = container.Controls.Add("Forms.ComboBox.1")
    cbo.Style = 2                           ' fmStyleDropDownList
    cbo.Left = X_COL: cbo.Top = y: cbo.Width = X_LINES - X_COL - 4: cbo.Height = CTL_H
    cbo.ListWidth = 260
    cbo.ListRows = 16
    cbo.List = Array(SRC_NOCOL)
    cbo.ListIndex = 0

    Set txtLines = container.Controls.Add("Forms.TextBox.1")
    txtLines.Left = X_LINES: txtLines.Top = y: txtLines.Width = X_DEF - X_LINES - 4: txtLines.Height = CTL_H
    txtLines.ControlTipText = "Zeile(n) innerhalb der Zelle: leer = alle, 1 = erste, 2-3, 3- = ab der dritten"

    Set txtDef = container.Controls.Add("Forms.TextBox.1")
    txtDef.Left = X_DEF: txtDef.Top = y: txtDef.Width = W_DEF: txtDef.Height = CTL_H
    txtDef.ControlTipText = "Standardwert, wenn die Zelle leer oder keine Spalte zugeordnet ist"
End Sub

Public Function ColIndex() As Long
    ColIndex = cbo.ListIndex
    If ColIndex < 0 Then ColIndex = 0
End Function

Public Function LineSpec() As String
    LineSpec = Trim$(txtLines.Text)
End Function

Public Function DefValue() As String
    DefValue = Trim$(txtDef.Text)
End Function

Public Sub SetColumns(ByVal items As Variant)
    Dim k As Long
    k = cbo.ListIndex
    mBusy = True
    cbo.List = items
    If k >= 0 And k <= UBound(items) Then cbo.ListIndex = k Else cbo.ListIndex = 0
    mBusy = False
End Sub

Public Sub SetColumn(ByVal col As Long)
    mBusy = True
    If col >= 0 And col < cbo.ListCount Then cbo.ListIndex = col Else cbo.ListIndex = 0
    mBusy = False
End Sub

Public Sub SetValues(ByVal col As Long, ByVal spec As String, ByVal sDef As String)
    SetColumn col
    mBusy = True
    txtLines.Text = spec
    txtDef.Text = sDef
    mBusy = False
End Sub

Private Sub Changed()
    If mBusy Then Exit Sub
    If Owner.Loading Then Exit Sub
    Owner.OnMapChanged
End Sub

Private Sub cbo_Change()
    Changed
End Sub

Private Sub txtLines_Change()
    Changed
End Sub

Private Sub txtDef_Change()
    Changed
End Sub

