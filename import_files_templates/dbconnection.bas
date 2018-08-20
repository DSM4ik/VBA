Attribute VB_Name = "dbconnection"
Option Explicit
Public clsCn As New clsConn
Public Const cnstPWD = "Hjccbz2017"


Function ADODBConnectionOracle(ByRef cn As ADODB.Connection, Optional ByVal adCursorLocation As Long = adUseServer, Optional ByVal adMode As Long = adModeShareDenyNone) As Boolean
 ' устанавливает соединение для Jet-провайдера
 Dim cntTime As Date
 Dim PWD As String
 
 '  PWD = InputBox("Введите пароль")
 ' -- If PWD = "" Then End
         
     cntTime = Now
         'установить обработчик ошибок:
     On Error GoTo err_not_connection
         'инициализировать cn
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
         
                  
         ADODBConnectionOracle = True 'возвращаемое значение
                        
         Exit Function 'обойти обработчик ошибок
         
err_not_connection:
         MsgBox "Не могу открыть БД для чтения." & Chr(13) & _
                 "Необходимо обратиться к администратору БД!!!", vbCritical, "СТОП"
         ADODBConnectionOracle = False
     Exit Function 'обойти обработчик ошибок



         
 End Function


Function ADODBConnectionAcc(ByRef cn As ADODB.Connection, ByVal dbPath As String, ByVal dbName As String, _
                         Optional ByVal adCursorLocation As Long = adUseServer, _
                         Optional ByVal adMode As Long = adModeRead) As Boolean
 ' устанавливает соединение для Jet-провайдера

         'установить обработчик ошибок:
         On Error GoTo err_not_connection
         If Dir(dbPath & dbName) = "" Then GoTo err_not_connection
         'инициализировать cn
         Set cn = CreateObject("ADODB.Connection")
         cn.Provider = "Microsoft.ACE.OLEDB.12.0"
         cn.ConnectionString = dbPath & dbName
         cn.CursorLocation = adCursorLocation
         cn.ConnectionTimeout = 100
'Блокировка должна иметь значение
         cn.Mode = adMode '
         cn.Open , , , 16
         Do While cn.State <> 1
          DoEvents
         Loop
         ADODBConnectionAcc = True 'возвращаемое значение
         Exit Function 'обойти обработчик ошибок
err_not_connection:
         ADODBConnectionAcc = False
         MsgBox "Не могу открыть - " & dbPath & dbName & " для чтения." & Chr(13) & _
                "При длительном появлении данного сообщения (более 15мин.) Вам необходимо обратиться к администратору БД!!", vbCritical, "СТОП"
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
        .LockType = adLockType    'adLockBatchOptimistic записи блокируются только в момент записи в базу данных
        .CacheSize = 1
        .Open , , , , 16
        Do While .State <> 1
          DoEvents
        Loop
        
End With
 Exit Function 'обойти обработчик ошибок
 
err_not_rst: 'Err.Description
   MsgBox Err.Description & "." & Chr(13) & Err.Source, vbCritical, "СТОП"
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
        .LockType = adLockType    'adLockBatchOptimistic записи блокируются только в момент записи в базу данных
        .CacheSize = 1
        .Open , , , , 16
        Do While .State <> 1
          DoEvents
        Loop
        
End With
 Exit Function 'обойти обработчик ошибок
 
err_not_rst: 'Err.Description
   MsgBox Err.Description & "." & Chr(13) & Err.Source, vbCritical, "СТОП"
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
    MsgBox "Текст ошибки - " & Err.Description & vbNewLine & vbNewLine & " Инструкция - " & strSQL, vbCritical, Err.Number & " - " & Err.Description
  End

End Function



