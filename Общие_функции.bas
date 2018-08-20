Attribute VB_Name = "Общие_функции"
Option Explicit


Function ConvertToDate()
Dim i As Long
Dim cnt As String
  For i = 1 To Len(cnt)
    If IsNumeric(Mid(cnt, i, 1)) Then
     End If
  Next
  

End Function

Function ArrFromRst(ByVal rst As ADODB.Recordset, ByRef arr)
Dim i As Long, J As Long, fldCount As Long, cntRecCount As Long

fldCount = rst.Fields.Count - 1: cntRecCount = 0

If Not rst.EOF Or Not rst.BOF Then
     Do Until rst.EOF
         cntRecCount = cntRecCount + 1
     rst.MoveNext
     Loop
   rst.MoveFirst
End If


ReDim arr(fldCount, cntRecCount - IIf(cntRecCount > 0, 1, 0))
    
    Do Until rst.EOF
      For J = 0 To fldCount
         arr(J, i) = rst.Fields(J)
      Next J
    
    rst.MoveNext
    i = i + 1
    Loop

End Function



Public Function DateSQLAccess(ByVal cntDate As Date) As String
    DateSQLAccess = CStr("#" & Month(cntDate) & "/" & Day(cntDate) & "/" & Year(cntDate) & "#")
 End Function

Public Function DateSQLServer(ByVal cntDate As Date) As String
    DateSQLServer = CStr(Month(cntDate) & "/" & Day(cntDate) & "/" & Year(cntDate))
End Function

Public Function DateSQLOracle(ByVal cntDate As Date) As String
    DateSQLOracle = "'" & CStr(Day(cntDate) & "/" & Month(cntDate) & "/" & Year(cntDate)) & "'"
End Function


 Function Get_User_Name() As String
     Get_User_Name = Environ("USERNAME")
 End Function

Function Get_User_FullName() As String
    With GetObject("LDAP://" & CreateObject("ADSystemInfo").UserName)
        Get_User_FullName = .FullName
    End With
End Function


Function ArrayPeriod(ByRef myDateARR, ByVal DateBeg As Date, ByVal DateEnd As Date, Optional ByVal TypeLoad As String = "Выгрузка Помесячная")

Dim i As Integer, J As Integer, cntDateOtchet As Date, T As Integer
' определяем матрицу периодов отчетности
  
    If TypeLoad = "Выгрузка Помесячная" Then
            ReDim myDateARR(0)
            myDateARR(0) = CDate(DateBeg)
            ' создаем матрицу дат отчетности выбранного периода
                  For J = Year(DateBeg) To Year(DateEnd)
                      For i = 1 To 12
                       cntDateOtchet = CDate("01." & i & "." & J)
                         If (cntDateOtchet > DateBeg And DateEnd > DateBeg And cntDateOtchet <= DateEnd) Then
                            T = T + 1
                            ReDim Preserve myDateARR(T)
                            myDateARR(T) = cntDateOtchet
                         End If
                      Next i
                  Next J
            If Day(DateEnd) <> 1 Or DateBeg = DateEnd Then
                     ReDim Preserve myDateARR(T + 1)
                     myDateARR(T + 1) = CDate(DateEnd)
            End If
                  
           If Day(DateEnd) <> 1 Or DateBeg = DateEnd Then
                     ReDim Preserve myDateARR(T + 1)
                     myDateARR(T + 1) = CDate(DateEnd)
           End If
    Else
       myDateARR = Array(CDate(DateBeg), CDate(DateEnd))
    
    End If
End Function





Function ShablonCopyPaste(ByVal shShablon As Worksheet, ByVal shRes As Worksheet, ByVal NameProc As String, _
                          Optional ByVal cntRowBeg As Long = 1, Optional ByVal cntClmnBeg As Long = 1)
Dim rngShablon As Range, RowEnd As Long, clmnEnd As Long, i As Long
' копируем шаблон в нужное нам место.

shRes.AutoFilterMode = False
shShablon.AutoFilterMode = False
shShablon.UsedRange.EntireRow.Hidden = False: shShablon.UsedRange.EntireColumn.Hidden = False

Clear_Sheet shRes.Cells, xlThemeColorDark1, -0.499984740745262
Set rngShablon = shShablon.Rows("1:" & iRowEnd(shShablon))
rngShablon.Copy Destination:=shRes.Cells(cntRowBeg + 2, cntClmnBeg)
RowEnd = iRowEnd(shRes): clmnEnd = iColumnEnd(shRes)
For i = cntClmnBeg To clmnEnd
        shRes.Columns(i).ColumnWidth = shShablon.Columns(i).ColumnWidth
Next
shRes.Columns(clmnEnd).Delete
GroupRow_On shRes, RowEnd
GroupClmn_On shRes, clmnEnd

End Function

Function Group_OFF(ByVal rng As Range)
' снятие всех группировок и показ всех скрытых строк и колонок
With rng
            .EntireRow.OutlineLevel = 1
            .EntireColumn.OutlineLevel = 1
            .EntireRow.Hidden = False
            .EntireColumn.Hidden = False
End With

End Function

Function GroupRow_On(ByVal shRes As Worksheet, ByVal RowEnd As Long)
' включение группировока строк
Dim i As Long
For i = RowEnd To 1 Step -1
     If shRes.Cells(i, 1).EntireRow.OutlineLevel > 1 And Not shRes.Cells(i, 1).EntireRow.Hidden Then _
                      shRes.Cells(i, 1).EntireRow.Hidden = True
         
Next

End Function

Function GroupClmn_On(ByVal shRes As Worksheet, ByVal clmnEnd As Long)
' включение группировока колонок
Dim i As Long
For i = clmnEnd To 1 Step -1
     If shRes.Cells(i, 1).EntireColumn.OutlineLevel > 1 And Not shRes.Cells(i, 1).EntireColumn.Hidden Then _
                      shRes.Cells(i, 1).EntireColumn.Hidden = True
         
Next

End Function



Function SetClmRange(ByVal sh As Worksheet, ByVal nmColumn As String, ByVal NameProc As String) As Range
Dim rng As Range
Set rng = FindElement(sh, nmColumn, nmColumn, NameProc)
Set SetClmRange = sh.Range(rng.Address & ":" & sh.Cells(iRowEnd(sh), rng.Column).Address)

End Function

Function ArrFiltr(ByVal shIstFiltr As Worksheet, ByVal nmClmnIstFiltr As String, ByRef arr)
'формируем сортированную матрицу
Dim cntRow As Long, cntClmn As Long, i As Long, J As Long, N As Long
Dim rngFiltr As Range, cnt As String, arrSh, Flag As Boolean

Set rngFiltr = FindElemArr(shIstFiltr, nmClmnIstFiltr, cntRow, cntClmn, "FiltrDop")
arrSh = shIstFiltr.Range(shIstFiltr.Cells(cntRow + 1, cntClmn).Address & ":" & shIstFiltr.Cells(iRowEnd(shIstFiltr), cntClmn).Address).Value

ReDim arr(0)
' формируем матрицу уникальных записей
For i = 1 To UBound(arrSh, 1)
    Flag = True
    cnt = arrSh(i, 1)
        If cnt <> "" Then
                  For J = 0 To UBound(arr)
                     Flag = True
                     If cnt = arr(J) Then Exit For
                     Flag = False
                  Next J
          
                  If Not Flag Then
                        ReDim Preserve arr(N): arr(N) = cnt: N = N + 1
                  End If
        End If
Next i

        'сортируем матрицу по возрастанию
        For i = 0 To UBound(arr)
                   cnt = arr(i)
          For J = i + 1 To UBound(arr)
            If cnt > arr(J) Then
                                     arr(i) = arr(J)
                                     arr(J) = cnt
                                     cnt = arr(i)
            End If
          Next J
        Next

End Function

Function FindElement(ByVal sh As Worksheet, ByVal cntNameRow As String, ByVal cntNameClmn As String, ByVal NamePoisk As String) As Range
' опредеяем значение в ячейке по названию колонки и строки
Dim ClmnID As Long, cntRow As Long, cntClmn As Long
 
 FindElemArr sh, cntNameClmn, cntRow, cntClmn, NamePoisk:     ClmnID = cntClmn
 FindElemArr sh, cntNameRow, cntRow, cntClmn, NamePoisk
 Set FindElement = sh.Cells(cntRow, ClmnID)


End Function

Function FindElemArr(ByVal shIst As Worksheet, ByVal cntElement, ByRef cntRow As Long, _
                              ByRef cntClmn As Long, ByVal cntRes As String, Optional ByVal FlagAdd As Boolean = False) As Range
Dim i As Long, J As Long, N As Integer, arrIst
'ищем элемент в матрице размерности (M,N) - на заданном листе по значению ячейки'

arrIst = shIst.Range("$A$1:" & shIst.Cells(iRowEnd(shIst), iColumnEnd(shIst)).Address).Value

For J = 1 To UBound(arrIst, 1)
  For i = 1 To UBound(arrIst, 2)
    If Trim(CStr(arrIst(J, i))) = Trim(CStr(cntElement)) Then
            cntRow = J
            cntClmn = i
            N = N + 1
    End If
 Next i
Next J

    If N = 0 And Not FlagAdd Then
        MsgBox "В файле источнике - '" & shIst.Parent.Name & "', на листе - '" & shIst.Name & _
              "' не обнаружен элемент - '" & cntElement & "'!!!", vbCritical, "ПРОГРАММА СТОП - " & cntRes
        End
    End If
    
    If N > 2 Then
       MsgBox "В файле источнике - '" & shIst.Parent.Name & "', на листе - '" & shIst.Name & _
              "' обнаружен более одного раза элемент - '" & cntElement & "'!!!", vbCritical, "ПРОГРАММА СТОП"
       End
    End If
    
 Set FindElemArr = Nothing
 If N = 1 Then Set FindElemArr = shIst.Cells(cntRow, cntClmn)
    
End Function

Public Function AnalizeFlagBSPL(ByVal cntTypeOtchet As String) As String

 Select Case cntTypeOtchet
      Case "Балансовые Остатки", "Балансовые Остатки RUR", "Балансовые Остатки Валюта"
            AnalizeFlagBSPL = "BS"
      Case "Среднедневые Балансовые Остатки", "Среднедневые Балансовые Остатки RUR", "Среднедневые Балансовые Остатки Валюта"
            AnalizeFlagBSPL = "AvrBS"
      Case "Отчет о прибылях и убытках"
            AnalizeFlagBSPL = "PL"
 End Select
End Function

Public Function iRowEnd(ByVal wSheet As Worksheet)
Dim RowEnd As Object
Set RowEnd = wSheet.Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
       If RowEnd Is Nothing Then
                      iRowEnd = 1
                      Else: iRowEnd = RowEnd.Row
       End If

End Function

Public Function iColumnEnd(ByVal Sheet As Worksheet)
Dim ColumnEnd As Object
Set ColumnEnd = Sheet.Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns)
       If ColumnEnd Is Nothing Then
                      iColumnEnd = 1
                      Else: iColumnEnd = ColumnEnd.Column
       End If

End Function

Function GetExcelWorkbook() As Excel.Workbook
Dim objExl As Excel.Application

    Set objExl = CreateObject("Excel.Application")
    objExl.DisplayAlerts = False
    objExl.DefaultSaveFormat = xlExcel12
    objExl.Visible = False
Set GetExcelWorkbook = objExl.Workbooks.Add

End Function

Function GetExcelWorksheet(ByVal wb As Excel.Workbook) As Excel.Worksheet

Set GetExcelWorksheet = wb.Worksheets.Add
    GetExcelWorksheet.Cells.Font.Name = "Times New Roman"
    GetExcelWorksheet.Cells.Font.Size = 12

End Function

Function RaskraskaZagolovok(ByVal rst As ADODB.Recordset, ByVal sh As Excel.Worksheet, ByVal PeriodOtchet As String, ByVal NameOtchet As String)

Dim i As Integer

 
PutFields_To_Excel sh, rst, 3, 1, xlThemeColorDark1, -0.149998474074526
sh.Cells(1, 1) = "Наименование отчета - " & NameOtchet
sh.Cells(2, 1) = "Период отчета - " & PeriodOtchet

End Function

Function MakeStrCritery(ByVal tblSaldoResult As String, ByVal cntID_DIVISION As String, _
                        ByRef cntCritery_ID_CLIENT As String, ByRef cntCritery_NUM_ACCOUNT As String, _
                        ByRef cntCritery_KodBS As String, ByRef cntCritery_KodPL As String, ByRef cntCritery_ID_DIVISION As String)

'Формируем строки критериев для отбора по маскам введенным в форму
Dim arr, arrIDClient As String, arrMask_NUM_ACCOUNT As String, LenMask As Integer, arrKodBS As String, _
    arrKodPL As String, arrID_DIVISION As String
Dim i As Integer

'разбираем введенные строки на элементы для SQL запроса. Формируем критерии отбора информации
    arrIDClient = arrStrCritery(cntCritery_ID_CLIENT)
    arrMask_NUM_ACCOUNT = arrStrCritery(cntCritery_NUM_ACCOUNT, "'"):   arrKodBS = arrStrCritery(cntCritery_KodBS, "'")
    arrKodPL = arrStrCritery(cntCritery_KodPL, "'"):                       arrID_DIVISION = arrStrCritery(cntID_DIVISION)
    
    cntCritery_ID_CLIENT = "":  cntCritery_NUM_ACCOUNT = "": cntCritery_KodBS = ""
    
    If Len(arrIDClient) > 0 Then cntCritery_ID_CLIENT = tblSaldoResult & ".ID_CLIENT in (" & arrIDClient & ")"
    If Len(arrID_DIVISION) > 0 Then cntCritery_ID_DIVISION = tblSaldoResult & ".ID_DIVISION in (" & arrID_DIVISION & ")"
    
    If Len(arrMask_NUM_ACCOUNT) > 0 Then
       arr = Split(arrMask_NUM_ACCOUNT, ",", -1)
       cntCritery_NUM_ACCOUNT = "(" & tblSaldoResult & ".NUM_ACCOUNT like " & arr(0)
      For i = 1 To UBound(arr)
        cntCritery_NUM_ACCOUNT = cntCritery_NUM_ACCOUNT & " OR " & tblSaldoResult & ".NUM_ACCOUNT like " & arr(i)
      Next
        cntCritery_NUM_ACCOUNT = cntCritery_NUM_ACCOUNT & ")"
    End If
    
    If Len(arrKodBS) > 0 Then cntCritery_KodBS = tblSaldoResult & ".[Код статьи BS] in (" & arrKodBS & ")"
    If Len(arrKodPL) > 0 Then cntCritery_KodPL = tblSaldoResult & ".[Код статьи PL] in (" & arrKodPL & ")"
    

End Function




Function arrStrCritery(ByVal cntStr As String, Optional ByVal chk As String = "") As String
Dim i As Integer, ChrI As String, arrStr As String

arrStr = ""
For i = 1 To Len(cntStr)
    ChrI = Mid(cntStr, i, 1)
    If Len(arrStr) = 0 And (IsNumeric(ChrI) Or ChrI = "%") Then arrStr = chk
    If IsNumeric(ChrI) Or ChrI = "%" Then arrStr = arrStr & ChrI
    
    If i < Len(cntStr) Then
      If Len(arrStr) > 0 And Not IsNumeric(ChrI) And ChrI <> "%" And _
         (IsNumeric(Mid(cntStr, i + 1, 1)) Or Mid(cntStr, i + 1, 1) = "%") And i > 0 Then arrStr = arrStr & chk & "," & chk
    Else
        If Len(arrStr) > 0 Then arrStr = arrStr & chk
    End If
Next
      arrStrCritery = arrStr

End Function


Function RaskrasRange(ByVal rng As Excel.Range, Optional ByVal Fon As Long = xlThemeColorDark1, _
                                                Optional ByVal TintShade As Double = -0.149998474074526)
With rng
  .Interior.ThemeColor = Fon
  .Interior.TintAndShade = TintShade
        
   Border rng
  .Font.Bold = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With

End Function

Function Clear_Sheet(ByVal rng As Range, Optional ByVal cntThemeColor As Double = xlThemeColorDark1, _
                                         Optional ByVal cntInerior As Double = -0.499984740745262)

Application.ScreenUpdating = False

rng.Clear
rng.Interior.ThemeColor = cntThemeColor
rng.Interior.TintAndShade = cntInerior


End Function


Function Border(ByVal rng As Excel.Range)

With rng
   With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous:        .ColorIndex = 0:        .TintAndShade = 0:        .Weight = xlThin
    End With
    
 End With
 
 End Function
 
 
Function PutOtchetToExcel(ByVal rst As ADODB.Recordset, ByVal sh As Worksheet, ByVal N_record As Long, _
                          ByVal arrNameParam, ByVal ArrParam, Optional ByVal NameOtchet As String = "Отчет")
'заголовок отчета с параметрами
Dim i As Integer
With sh
  
  'заголовок отчета с параметрами
  For i = 0 To UBound(arrNameParam)
    .Cells(i + 1, 1) = arrNameParam(i) & " = " & ArrParam(i)
  Next
  
 
 ' проверяем объем отбираемых записей. Останвливаем прогу, если запаисей более 1048000
    If Not Export_Excel(rst, sh, N_record, UBound(arrNameParam) + 2, 1, xlThemeColorLight2, 0.799981688894314) Then
         sh.Parent.Parent.Quit
         rst.Close: rst.ActiveConnection.Close
         End
    End If
   .Name = NameOtchet:    .Visible = True:    .Activate:     ActiveWindow.ScrollRow = 1
      
End With
End Function






Function Export_Excel(ByVal rst As ADODB.Recordset, ByVal sh As Worksheet, ByVal N_Records As Long, _
                      Optional ByVal RowBeg As Long = 1, Optional ByVal clmnBeg As Long = 1, _
                      Optional ByVal Fon As Long = xlThemeColorDark1, _
                      Optional ByVal clr As Double = -0.149998474074526, _
                      Optional ByVal FlagZagolovok As Boolean = True) As Boolean
'экспорт анализирует колличество выгружаемых записей и останавливает работу при ьольшом объеме
Const R_Copy = 5000   ' колличество строк в единоразовой выгрузке

Dim fld   As ADODB.Field, N As Long, RowEnd As Long, cntN As Long
Dim rst_Cl As ADODB.Recordset, pi As ProgressIndicator, FlagZapis As Boolean, arr

'проверяем возможность не выйти за пределы листа
If iRowEnd(sh) + 1 + N_Records > 1045000 Then GoTo errln2

Export_Excel = True
     
      If FlagZagolovok Then PutFields_To_Excel sh, rst, RowBeg, clmnBeg, Fon, clr
      If N_Records < 100001 Then
        sh.Cells(RowBeg + 1, clmnBeg).CopyFromRecordset rst
        Exit Function
      End If
      
Set pi = New ProgressIndicator:  pi.Show "Выгрузка данных в колличетсве - " & N_Records
                   pi.ShowPercents = True:   pi.StartNewAction 0, 100, , , , 0
                   
' создаем структуру отвязанного рекордсета по образцу rst
 Dim i As Long

FlagZapis = True

'ReDim Arr(1 To R_Copy, 1 To rst.Fields.Count)
Do Until rst.EOF
   If FlagZapis Then
         pi.UpdateSingleAction (cntN + 1) / (N_Records + 1) * 100, , , "Выгружено строк - " & R_Copy * N
         FlagZapis = False
   End If
       
         sh.Cells(iRowEnd(sh) + 1, clmnBeg).CopyFromRecordset rst, R_Copy
         N = N + 1
         
         'копируем часть данных в рекордсет
                  
        '          For I = 0 To rst.Fields.Count - 1
         '           Arr(N, I + 1) = rst.Fields(I)
          '        Next
            
         ' если число строк в рекордсете более требуемой величины R_count, то экспортируем данные
   '      If N = R_Copy Or rst.EOF Then
    '        RowEnd = iRowEnd(sh) + 1
       '     N = 0: FlagZapis = True
        '    sh.Cells(RowEnd, clmnBeg).Resize(UBound(Arr, 1), UBound(Arr, 2)).NumberFormat = "@"
         '   sh.Cells(RowEnd, clmnBeg).Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
          '  Arr = 0
           ' ReDim Arr(1 To R_Copy, 1 To rst.Fields.Count): DoEvents
      '   End If
 '  rst.MoveNext
Loop
    pi.Hide
    Exit Function

errln2:
   MsgBox "Недостаточно ресурсов для выгрузки данных. Уменьшите запрос, закройте лишние приложения!!!"
   Export_Excel = False
   
End Function


Function MakeArrFrom_RST(ByVal rst As ADODB.Recordset, ByRef arrSum, ByRef arrKod, ByVal FldSum As String, ByVal fldKod As String) As Currency
Dim i As Long, J As Long, N As Long, T As Long, FlagKod As Boolean
'ФОРМИРУЕМ МАТРИЦУ сумм ИЗ РЕКОРДСЕТА
    ReDim Preserve arrSum(1, UBound(arrKod))
    If Not rst.EOF Or Not rst.BOF Then rst.MoveFirst
     Do Until rst.EOF
            FlagKod = False
            T = IIf(rst.Fields("DopTask") = "Основное", 0, 1)
            For J = 0 To UBound(arrKod)
               If arrKod(J) = rst.Fields(fldKod).Value Then
                  N = J
                  FlagKod = True
                  Exit For
               End If
            Next
               If Not FlagKod Then
                 N = IIf(IsEmpty(arrKod(0)), 0, J)
                 ReDim Preserve arrSum(1, N):      ReDim Preserve arrKod(N)
               End If
              
               arrKod(N) = rst.Fields(fldKod).Value
               arrSum(T, N) = arrSum(T, N) + IIf(IsNull(rst.Fields(FldSum).Value), 0, rst.Fields(FldSum).Value)
               
        rst.MoveNext
     Loop



End Function




Function PutFields_To_Excel(ByVal sh As Worksheet, ByVal rst As ADODB.Recordset, ByVal RowBeg As Long, ByVal clmnBeg As Long, _
                            Optional ByVal Fon As Long = xlThemeColorDark1, Optional ByVal clr As Double = -0.149998474074526)

Dim i As Long, rng As Range

With sh
For i = 0 To rst.Fields.Count - 1
     .Cells(RowBeg, clmnBeg + i) = rst.Fields(i).Name
     If IsDate(rst.Fields(i)) Then .Columns(clmnBeg + i).NumberFormat = "dd/mm/yyyy"
Next
Set rng = .Range(.Cells(RowBeg, clmnBeg).Address & ":" & .Cells(RowBeg, clmnBeg + rst.Fields.Count - 1).Address)
  RaskrasRange rng, Fon, clr
  rng.Columns.AutoFit
End With
End Function

Function OgranichenieRecords_EndProgramm(ByRef cnDateSQL As ADODB.Connection, ByVal tbl As String) As Long
 Dim rst As ADODB.Recordset
  
 Set rst = Rst_Conn(cnDateSQL, "SELECT COUNT(*) as Cnt FROM " & tbl, adUseServer)
 OgranichenieRecords_EndProgramm = rst.Fields("Cnt")
 If OgranichenieRecords_EndProgramm > 1047999 Then
    MsgBox "Не хватает памяти для выгрузки данных." & Chr(13) & _
           "Ограничте Вашу выборку или закройте лишние приложения!!!" & Chr(13) & _
           "Было отобрано колличество записей - '" & Format(OgranichenieRecords_EndProgramm, "#,##0") & "'", vbCritical
       cnDateSQL.Close
    End
 End If
End Function

Function Arr_Segment(ByVal cn As ADODB.Connection, ByVal tblPlane As String, ByVal fldKod_BSPL As String, ByVal fldSegment As String, _
                             ByVal arrKod, ByVal cntIDSegment As String, ByRef arrSegment)
' по таблице библиотек кодов  определяем сегмент для планирования
    Dim rst As ADODB.Recordset, i As Long, strSQL As String
    
    ReDim arrSegment(UBound(arrKod, 1))
   
If cntIDSegment = "" Then
   For i = 1 To UBound(arrKod, 1)
     strSQL = "SELECT distinct [" & fldSegment & "] FROM " & tblPlane & " WHERE [" & fldKod_BSPL & "]='" & arrKod(i, 1) & "'"
     Set rst = Rst_Conn(cn, strSQL)
       If rst.RecordCount = 1 Then
                arrSegment(i) = rst.Fields(fldSegment)
       Else: arrSegment(i) = "АНАЛИТ"
       End If
   Next
Else
     For i = 0 To UBound(arrKod, 1)
                arrSegment(i) = cntIDSegment
     Next
End If
End Function









 
 
Function Rst_FreeStructure(ByVal rst As ADODB.Recordset) As ADODB.Recordset
    ' создаем структуру отвязанного рекордсета по образцу
    Dim fld  As ADODB.Field
    
    Set Rst_FreeStructure = New ADODB.Recordset
    For Each fld In rst.Fields
          Rst_FreeStructure.Fields.Append fld.Name, fld.Type, fld.DefinedSize, adFldIsNullable
    Next
    Rst_FreeStructure.Open , , adOpenDynamic, adLockOptimistic
End Function

Function FlagSelection(ByVal cmbObj As ComboBox, Optional ByRef i As Long = 0) As Boolean
 'данная функция работает для фильтра с уникальными названиеми в колонке 0

   FlagSelection = False
   For i = 0 To cmbObj.ListCount - 1
       If cmbObj.Text = cmbObj.List(i, 0) Then
              cmbObj.Value = cmbObj.Text
              FlagSelection = True
              Exit For
       End If
   Next
End Function

Sub Arr_ListBox(rst As ADODB.Recordset, ByVal cntFiltr As Object, ByVal arr, ByVal cntClmnWidth As String, _
                Optional ByVal FlagClear As Boolean = True)
Dim i As Long, J As Long
 With cntFiltr
        If FlagClear Then
                     .Clear: .Value = ""
        End If
        .ColumnCount = UBound(arr) + 1
        .ColumnWidths = cntClmnWidth
        .TextAlign = fmTextAlignLeft
 
             i = .ListCount
             If Not rst.EOF Or Not rst.BOF Then
                      rst.MoveFirst
                               Do Until rst.EOF
                                  
                                  .AddItem
                                  For J = 0 To UBound(arr)
                                   cntFiltr.List(i, J) = IIf(IsNull(rst.Fields(arr(J)).Value), "", rst.Fields(arr(J)).Value)
                                   If IsDate(rst.Fields(arr(J))) And cntFiltr.List(i, J) <> "" And rst.Fields(arr(J)).Name <> "Number" Then _
                                                cntFiltr.List(i, J) = Format(rst.Fields(arr(J)).Value, "dd.mm.yyyy")
                                  Next
                                  i = i + 1
                                rst.MoveNext
                               Loop
             End If
 End With
End Sub



Sub Check_On_Off_Click(ByVal cntListBox As Object)
'включение/отключение отбора на листбоксе
Dim i As Long, FlagSelection As Boolean
     With cntListBox
         FlagSelection = False
            For i = 0 To .ListCount - 1
                     If .Selected(i) = True Then
                                  FlagSelection = True:    Exit For
                     End If
            Next
         
         .Selected(0) = False
         For i = 0 To .ListCount - 1
               If FlagSelection Then
                         .Selected(i) = False
               Else:     .Selected(i) = True
               End If
         Next
     End With
End Sub

Function Selecttion_ON(ByVal objFiltr As Object)
Dim i As Long
For i = 0 To objFiltr.ListCount - 1
    objFiltr.Selected(i) = True
Next
End Function

Function Selecttion_Off(ByVal objFiltr As Object)
Dim i As Long
For i = 0 To objFiltr.ListCount - 1
    objFiltr.Selected(i) = False
Next
End Function




