VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module      : ProgressIndicator
' Автор     : EducatedFool  (Игорь)                    Дата: 23.10.2011
' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
'---------------------------------------------------------------------------------------

Public FP As New F_Progress

Private SubActionIndex As Single, SubActionsCount As Single
Private FPVisible As Boolean, FPStartTime As Date, Position As Integer
Private PrS As Integer, PrE As Integer, Percent As Double

Public Parent As ProgressIndicator
Public ShowPercents As Boolean, ShowTime As Boolean

Function AddChildIndicator(ByVal Caption As String, Optional ByVal FPPosition As Integer = 1) As ProgressIndicator
    ' создаёт дочерний индикатор, и отображает его
    Set AddChildIndicator = New ProgressIndicator
    Set AddChildIndicator.Parent = Me
    AddChildIndicator.Show Caption, FPPosition
End Function

Private Sub Class_Initialize()
    ' параметры по умолчанию для вновь создаваемого индикатора
    Set FP = New F_Progress: ShowPercents = True: FPVisible = True
    PrS = 0: PrE = 100: Set_ProgressBar 0: FP.PrBar.Caption = ""
    FPStartTime = Now: ShowTime = True: ShowPercents = True
End Sub

Sub Show(ByVal Caption As String, Optional ByVal FPPosition As Integer = 0)
    ' отображает прогресс-бар
    SetProgressFormCaption Caption: On Error Resume Next:
    FP.PrBar.Width = ProgressBar_Default_Width
    Position = FPPosition
    FP.Tag = Caption: FP.Show:
    If Position <> 0 Then Move Position
    FP.Repaint: DoEvents
End Sub

Sub Hide()
 ' скрытие прогресс-бара
Unload FP
FPVisible = False
Application.ScreenUpdating = True
Application.ScreenUpdating = False
DoEvents
End Sub

Sub Repaint()
    FP.Repaint: DoEvents
End Sub

Sub Move(ByVal Position As Integer)    ' позиция прогресс-бара на экране по вертикали
    If Abs(Position) > 3 Then Exit Sub
    FP.Top = FP.Top + (FP.Height + 3) * Position
End Sub

Public Property Get Visible(): Visible = FPVisible: End Property




' установка заголовка формы и надписей на индикаторе
Public Property Let Line1(ByVal NewValue As String): FP.L1.Caption = ProcessLabel(NewValue): End Property
Public Property Let Line2(ByVal NewValue As String): FP.L2.Caption = ProcessLabel(NewValue): End Property
Public Property Let Line3(ByVal NewValue As String): FP.L3.Caption = ProcessLabel(NewValue): End Property
Public Property Let Caption(ByVal NewValue As String): SetProgressFormCaption NewValue: End Property

Public Sub SetProgressFormCaption(Optional ByVal Caption As String = "")
    ' устанавливает заголовок формы прогресс-бара
    ' с учёток настроек (отображение таймера и процента выполнения)
    If Len(Caption) > 0 Then FP.Tag = Caption
    txt = Trim(FP.Tag): If ShowPercents Then txt = Round(Percent, 0) & " %   " & txt
    dt = Format(Now - FPStartTime, "HH:NN:SS")
    If ShowTime Then txt = "( " & dt & " )    " & txt
    FP.Caption = txt
End Sub

Function TimeToFinish() As String
    If Percent < 15 Then Exit Function    ' сложно предсказать время, когда всё только начинается...
    dt = (Now - FPStartTime) * (100 - Percent) / Percent
    TimeToFinish = IIf(Minute(dt) > 0, Minute(dt) & " мин. ", "") & Second(dt) & " сек."
    If dt < TimeSerial(0, 0, 1) Then TimeToFinish = "менее секунды"
    TimeToFinish = "Осталось до завершения: " & TimeToFinish
End Function

Sub SetFocus()    ' делает форму прогресс-бара активной
    FP.Show 0: If Position <> 0 Then Move Position
End Sub

Private Sub UpdateLabels(Optional ByVal L1_txt$, Optional ByVal L2_txt$, Optional ByVal L3_txt$)
    ' обновляем надписи на прогресс-баре (выводит только непустые строки)
    If L1_txt$ <> "" Then FP.L1.Caption = ProcessLabel(L1_txt$)
    If L2_txt$ <> "" Then FP.L2.Caption = ProcessLabel(L2_txt$)
    If L3_txt$ <> "" Then FP.L3.Caption = ProcessLabel(L3_txt$)
End Sub

Function ProcessLabel(ByVal txt As String) As String
    ' заменяет ключевые слова в строке txt на значения параметров индикатора
    txt = Replace(txt, "$index", SubActionIndex)
    txt = Replace(txt, "$count", SubActionsCount)
    txt = Replace(txt, "$time", TimeToFinish)
    ProcessLabel = txt
End Function

Sub SubAction(Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = "", Optional ByVal L3_txt$ = "")
    ' запуск операции внутри основного действия
    SubActionIndex = SubActionIndex + 1
    If SubActionIndex > SubActionsCount Then SubActionIndex = SubActionsCount
    Percent = PrS + (PrE - PrS) * ((SubActionIndex - 1) / SubActionsCount)
    UpdateLabels L1_txt$, L2_txt$, L3_txt$
    Set_ProgressBar Percent: DoEvents
End Sub

Sub UpdateSingleAction(ByVal Percent As Double, Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = "", Optional ByVal L3_txt$ = "")
   UpdateLabels L1_txt$, L2_txt$, L3_txt$
   Set_ProgressBar Percent: DoEvents
End Sub

Sub StartNewAction(Optional ByVal Pr_Start As Integer = 0, Optional ByVal Pr_End As Integer = 100, _
                   Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = " ", Optional ByVal L3_txt$ = " ", _
                   Optional ByVal Actions_Count As Single = 0)
    ' запуск нового действия - на шкале индикатора от Pr_Start % до Pr_End %
    ' в переменной Actions_Count получает количество операций внутри действия
    PrS = Pr_Start: PrE = Pr_End: SubActionIndex = 0: SubActionsCount = Actions_Count
    UpdateLabels L1_txt$, L2_txt$, L3_txt$
    Set_ProgressBar PrS
End Sub

Sub UpdateFromChild(ByVal ChildPercent As Double)
    ' отображение изменений на родительской форме при изменениях на дочерней
    If SubActionsCount = 0 Or SubActionsCount = 1 Then
        Percent = PrS + (PrE - PrS) * (ChildPercent / 100)
    Else
        ' например, SubActionIndex = 3: SubActionsCount=10: PrS = 50: PrE = 100: ChildPercent=40
        ' результат д.б. в диапазоне от 60 до 65, а именно равен 62
        
        Percent = PrS + ((PrE - PrS) * (SubActionIndex - 1) / SubActionsCount) + _
                  (PrE - PrS) / SubActionsCount * (ChildPercent / 100)
    End If
    Set_ProgressBar Percent
End Sub

Sub Set_ProgressBar(ByVal NewPercent As Double)
    ' изменение ширины индикатора
    Percent = NewPercent
    If NewPercent > 100 Then Percent = 100
    If NewPercent < 0 Then Percent = 0
    FP.PrBar.Width = Int(Percent * ProgressBar_Default_Width / 100)
    SetProgressFormCaption
    If Not Parent Is Nothing Then Parent.UpdateFromChild Percent
    FP.Repaint
End Sub
Function GetCurrentProgress() As Long    ' возвращает текущий процент выполнения
    If FP.PrBar.Width = 0 Then Exit Function
    GetCurrentProgress = FP.PrBar.Width / ProgressBar_Default_Width * 100
End Function

Private Sub Class_Terminate()    ' уничтожение экземпляра класса
    Unload FP: FPVisible = False
End Sub

Function ProgressBar_Default_Width() As Double    ' установка размера полосы по размеру формы
   ProgressBar_Default_Width = FP.Width - 18
End Function




