Attribute VB_Name = "dbconnection"
Option Explicit
Public clsCn As New clsConn
Public Const cnstPWD = "Hjccbz2017"


Function ADODBConnectionOracle(ByRef cn As ADODB.Connection, Optional ByVal adCursorLocation As Long = adUseServer, Optional ByVal adMode As Long = adModeShareDenyNone) As Boolean
 ' ������������� ���������� ��� Jet-����������
 Dim cntTime As Date
 Dim PWD As String
 
 '  PWD = InputBox("������� ������")
 ' -- If PWD = "" Then End
         
     cntTime = Now
         '���������� ���������� ������:
     On Error GoTo err_not_connection
         '���������������� cn
         Set cn = CreateObject("ADODB.Connection")
         cn.ConnectionString = "Provider=MSDAORA.1;Persist Security Info=False;Data Source=xxi;User ID=BankReports;Password='" & cnstPWD & "'"
         cn.CommandTimeout = 200
         cn.CursorLocation = adCursorLocation
         cn.Mode = adMode
         cn.Open
         
         Do While cn.State <> 1
          If Now >= DateAdd("s", 10, CDate(cntTime)) Then Exit Do
          DoEvents
         Loop
         
                  
         ADODBConnectionOracle = True '������������ ��������
                        
         Exit Function '������ ���������� ������
         
err_not_connection:
         MsgBox "�� ���� ������� �� ��� ������." & Chr(13) & _
                 "���������� ���������� � �������������� ��!!!", vbCritical, "����"
         ADODBConnectionOracle = False
     Exit Function '������ ���������� ������



         
 End Function


Function ADODBConnectionAcc(ByRef cn As ADODB.Connection, ByVal dbPath As String, ByVal dbName As String, _
                         Optional ByVal adCursorLocation As Long = adUseServer, _
                         Optional ByVal adMode As Long = adModeRead) As Boolean
 ' ������������� ���������� ��� Jet-����������

         '���������� ���������� ������:
         On Error GoTo err_not_connection
         If Dir(dbPath & dbName) = "" Then GoTo err_not_connection
         '���������������� cn
         Set cn = CreateObject("ADODB.Connection")
         cn.Provider = "Microsoft.ACE.OLEDB.12.0"
         cn.ConnectionString = dbPath & dbName
         cn.CursorLocation = adCursorLocation
         cn.ConnectionTimeout = 100
'���������� ������ ����� ��������
         cn.Mode = adMode '
         cn.Open , , , 16
         Do While cn.State <> 1
          DoEvents
         Loop
         ADODBConnectionAcc = True '������������ ��������
         Exit Function '������ ���������� ������
err_not_connection:
         ADODBConnectionAcc = False
         MsgBox "�� ���� ������� - " & dbPath & dbName & " ��� ������." & Chr(13) & _
                "��� ���������� ��������� ������� ��������� (����� 15���.) ��� ���������� ���������� � �������������� ��!!", vbCritical, "����"
         ADODBConnectionAcc = False
 End Function
 
 
 
Sub ADODBConnection_Excel(ByRef cn As ADODB.Connection, ByVal wbFullName As String)
       
        
Set cn = New ADODB.Connection
With cn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .ConnectionString = "Data Source=" & wbFullName & ";Extended Properties=""Excel 12.0;HDR=YES"""
    .CursorLocation = adUseClient
    .Open , , , 16
    Do While .State <> 1
          DoEvents
    Loop
End With

End Sub

Function Rst_Conn(ByVal strSQL As String, _
                  Optional ByVal adCursorLocation As Long = adUseClient, _
                  Optional ByVal adLockType As Long = adLockReadOnly) As ADODB.Recordset
 
On Error GoTo err_not_rst

If clsCn Is Nothing Then Set clsCn = New clsConn
If clsCn.cnOra.State <> 1 Then clsCn = New clsConn

Set Rst_Conn = CreateObject("ADODB.Recordset")
With Rst_Conn
        .ActiveConnection = clsCn.cnOra
        .Source = strSQL
        .CursorLocation = adCursorLocation
        .CursorType = adOpenDynamic
        .LockType = adLockType    'adLockBatchOptimistic ������ ����������� ������ � ������ ������ � ���� ������
        .CacheSize = 1
        .Open , , , , 16
        Do While .State <> 1
          DoEvents
        Loop
        
End With
 Exit Function '������ ���������� ������
 
err_not_rst: 'Err.Description
   MsgBox Err.Description & "." & Chr(13) & Err.Source, vbCritical, "����"
  Set clsCn = Nothing
  End
End Function

Function Rst_Conn_Acc(ByVal cn As ADODB.Connection, ByVal strSQL As String, _
                       Optional ByVal adCursorLocation As Long = adUseClient, _
                       Optional ByVal adLockType As Long = adLockReadOnly) As ADODB.Recordset
 
On Error GoTo err_not_rst

Set Rst_Conn_Acc = CreateObject("ADODB.Recordset")
With Rst_Conn_Acc
        .ActiveConnection = cn
        .Source = strSQL
        .CursorLocation = adCursorLocation
        .CursorType = adOpenDynamic
        .LockType = adLockType    'adLockBatchOptimistic ������ ����������� ������ � ������ ������ � ���� ������
        .CacheSize = 1
        .Open , , , , 16
        Do While .State <> 1
          DoEvents
        Loop
        
End With
 Exit Function '������ ���������� ������
 
err_not_rst: 'Err.Description
   MsgBox Err.Description & "." & Chr(13) & Err.Source, vbCritical, "����"
  Set clsCn = Nothing
  End
End Function


Function ConnectExecute(ByVal strSQL As String)

On Error GoTo ErrExec:
If clsCn Is Nothing Then Set clsCn = New clsConn
If clsCn.cnOra.State <> 1 Then clsCn = New clsConn

 clsCn.cnOra.BeginTrans
 clsCn.cnOra.Execute strSQL, , 16: DoEvents
   Do While clsCn.cnOra.State <> 1
          DoEvents
   Loop
 clsCn.cnOra.CommitTrans
 
 Exit Function
ErrExec:
    MsgBox "����� ������ - " & Err.Description & vbNewLine & vbNewLine & " ���������� - " & strSQL, vbCritical, Err.Number & " - " & Err.Description
  End

End Function



