VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public cnOra As ADODB.Connection

Private Sub Class_Initialize()

If cnOra Is Nothing Then
    If Not ADODBConnectionOracle(cnOra, adUseServer) Then
       MsgBox "База BankReports недоступна!!! Обратитесь к администратору.", vbCritical
       End
    End If
End If
End Sub

Function cnOra_OpenRead()
   cnOra.Close: cnOra.Mode = adModeRead: cnOra.Open
End Function

Function cnOra_OpenReadWrite()
   cnOra.Close: cnOra.Mode = adModeReadWrite: cnOra.Open
End Function


Private Sub Class_Terminate()
Set cnOra = Nothing
End Sub



