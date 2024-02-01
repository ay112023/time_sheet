Attribute VB_Name = "Module1"
' Расчёт рабочих часов в табеле службы эксплуатации Электроучастка ЭУ "ГП "НПО ПХЗ"

' Глобальные переменные
Dim gEveningStartTime As Integer
Dim gEveningEndTime As Integer
Dim gNightStartTime As Integer
Dim gNightEndTime As Integer
Dim gDayShiftStartTime As Integer
Dim gDayShiftEndTime As Integer
Dim gNightShiftStartTime As Integer
Dim gNightShiftEndTime As Integer
Dim gPersonEveningHours As Integer
Dim gPersonNightHours As Integer
Dim gTimeSheetColStart As Integer
Dim gTimeSheetColEnd As Integer
Dim gSheetNumCol As Integer
Dim gSrvCol As Integer
Dim gSheetRowStart As Integer
Dim gEveningHoursCol As Integer
Dim gNightHoursCol As Integer
Dim gTotalHoursCol As Integer
Dim gTotalDaysCol As Integer
' Инициализация глобальных переменных:
' Начало, конец: вечернего, ночного времени, дневной, ночной смен
Public Sub Init()
 Dim ServiceSheet As Worksheet
 Set ServiceSheet = Sheets("Служебный")
 gEveningStartTime = ServiceSheet.Cells(2, 2)
 gEveningEndTime = ServiceSheet.Cells(2, 3)
 gNightStartTime = ServiceSheet.Cells(3, 2)
 gNightEndTime = ServiceSheet.Cells(3, 3)
 gDayShiftStartTime = ServiceSheet.Cells(4, 2)
 gDayShiftEndTime = ServiceSheet.Cells(4, 3)
 gNightShiftStartTime = ServiceSheet.Cells(5, 2)
 gNightShiftEndTime = ServiceSheet.Cells(5, 3)
 gTimeSheetColStart = 5
 gTimeSheetColEnd = 21
 gSheetNumCol = 4
 gSrvCol = 28
 gSheetRowStart = 1
 gTotalDaysCol = 22
 gTotalHoursCol = 23
 gEveningHoursCol = 24
 gNightHoursCol = 25
End Sub
' Расчёт вечерних часов за один рабочий день
Private Function EveningOpen_Hours(ByVal Shift As Byte, ByVal ShiftPart As Byte, ByVal ShiftLatency As Integer)
  
 Dim EveningHours As Integer
 Dim EveningLatency As Integer
   
    
   If Shift = 1 Then ' Вечерние часы дневной смены
       EveningHours = (gDayShiftStartTime + ShiftLatency) - gEveningStartTime
       If EveningHours < 0 Then
         EveningHours = 0
       End If
   ElseIf Shift = 2 Then ' Вечерние часы ночной смены :
      If ShiftPart = 1 Then   '  1-я часть ночной смены от начала до 24-00
         If gNightShiftStartTime > gEveningStartTime Then ' Ночная смена начинается после начала вечернего времени
           EveningHours = (gEveningEndTime - gNightShiftStartTime)
           If ShiftLatency < EveningHours Then
              EveningHours = ShiftLatency
           End If
        Else                                               ' Ночная смена начинается до начала вечернего времени
          EveningHours = ShiftLatency - (gEveningStartTime - gNightShiftStartTime)
          EveningLatency = (gEveningEndTime - gEveningStartTime)
          If EveningHours > EveningLatency Then
            EveningHours = EveningLatency
          End If
        End If
      ElseIf ShiftPart = 2 Then ' 2-я часть ночной смены: следующие сутки от 24-00 до конца
         EveningHours = 0       ' считаем, что ночная смена закончитсяч ДО вечера, поэтому - 0
      End If
   End If
   
 EveningOpen_Hours = EveningHours
End Function
' Расчёт ночных часов за один рабочий день
Private Function NightOpen_Hours(ByVal Shift As Byte, ByVal ShiftPart As Byte, ByVal ShiftLatency As Integer)
  
  Dim NightHours As Integer
  
  If Shift = 1 Then ' Ночные часы дневной смены = 0
    NightHours = 0
  ElseIf Shift = 2 Then ' ночные часы ночной смены :
    If ShiftPart = 1 Then ' 1-я часть ночной смены от начала до 24-00
       If (gNightShiftStartTime > gNightStartTime) And (gNightShiftStartTime < gNightEndTime) Then ' Ночная смена начинается после начала ночного времени
         NightHours = gNightEndTime - gNightShiftStartTime
          If ShiftLatency < NightHours Then
            NightHours = ShiftLatency
          End If
       ElseIf gNightShiftStartTime < gNightStartTime Then ' Ночная смена начинается до начала ночного времени
           NightHours = ShiftLatency - (gNightStartTime - gNightShiftStartTime)
           If NightHours > (24 - gNightStartTime) Then
             NightHours = 24 - gNightStartTime
           End If
       Else
            NightHours = 0
       End If
    ElseIf ShiftPart = 2 Then ' 2-я часть ночной смены: следующие сутки от 24-00 до конца
       If gNightShiftEndTime > gNightEndTime Then
           NightHours = gNightEndTime
            If ShiftLatency < NightHours Then
              NightHours = ShiftLatency
            End If
       Else
           NightHours = gNightShiftEndTime
       End If
    End If
  End If
  
  NightOpen_Hours = NightHours
End Function
' Расчёт вечерних, ночных, общих рабочих часов и рабочих дней по одному рабочему за месяц
Private Sub Calculate_Person(CurrentSheet As Worksheet, RowStart As Integer)
   Dim IndexCol, IndexRow As Integer
   Dim CurrentValue, NextValue, Divider As String
   Dim DividerPos As Integer
   Dim ShiftLatency As Integer
   Dim Shift, ShiftPart As Byte
   Dim EveningHours, NightHours, TotalHours, TotalDays As Integer
      
   Divider = "\"
   EveningHours = 0
   NightHours = 0
   TotalHours = 0
   ShiftPart = 0
   TotalDays = 0
     
   For IndexRow = RowStart To (RowStart + 1)
    For IndexCol = gTimeSheetColStart To gTimeSheetColEnd
      CurrentValue = CurrentSheet.Cells(IndexRow, IndexCol)
      CurrentValue = Trim(CurrentValue)
      
      If Len(CurrentValue) > 0 Then
               
        If (CurrentValue Like "##" + Divider + "#") Or (CurrentValue Like "#" + Divider + "#") Then
         
         DividerPos = InStr(1, CurrentValue, Divider)
         ShiftLatency = CInt(Mid(CurrentValue, 1, DividerPos - 1))
         Shift = CByte(Mid(CurrentValue, DividerPos + 1, 1))
         
         
         If Shift = 2 Then
         
           If ShiftLatency = 8 Or ShiftLatency = 7 Then
             ShiftPart = 2
           End If
         
           'Следующие сутки ночной смены: 2-я половина смены если в табеле снова "/2"
           'ShiftPart = ShiftPart + 1
           
           'If ShiftPart > 2 Then
           '   ShiftPart = 1
           'End If
           
           'Если в 1-й день месяца есть ночная смена
           'проверяем, не следующие ли это сутки ночной смены
           'для этого анализируем значение 2-го дня месяца
           
           'If IndexCol = gTimeSheetColStart Then
           '  NextValue = Trim(CurrentSheet.Cells(IndexRow, IndexCol + 1))
           '  If Not (NextValue Like "#" + Divider + "2") Then
           '    ShiftPart = 2
           '  End If
           'End If
           
          Else
           ShiftPart = 0
          End If
                  
         EveningHours = EveningHours + EveningOpen_Hours(Shift, ShiftPart, ShiftLatency)
         NightHours = NightHours + NightOpen_Hours(Shift, ShiftPart, ShiftLatency)
         TotalHours = TotalHours + ShiftLatency
         TotalDays = TotalDays + 1
        ElseIf CurrentValue Like "#" Then
          TotalHours = TotalHours + CInt(CurrentValue)
          TotalDays = TotalDays + 1
        Else
          ShiftPart = 0
        End If
      End If
    Next
   Next
   
   CurrentSheet.Cells(RowStart, gTotalDaysCol) = TotalDays
   CurrentSheet.Cells(RowStart, gTotalHoursCol) = TotalHours
   CurrentSheet.Cells(RowStart, gEveningHoursCol) = EveningHours
   CurrentSheet.Cells(RowStart, gNightHoursCol) = NightHours
End Sub
' Расчёт часов и дней по табелю
Public Sub CalculateAll()
Dim IndexRow As Integer
Dim CurrentSheet As Worksheet
Dim CurrentSheetNum, TempSheetNum As String

Set CurrentSheet = ActiveWorkbook.ActiveSheet

Call Init

IndexRow = gSheetRowStart
 Do
  TempSheetNum = Trim(CurrentSheet.Cells(IndexRow, gSheetNumCol))
  If TempSheetNum Like "#####" And CurrentSheetNum <> TempSheetNum Then
     CurrentSheetNum = TempSheetNum
     Call Calculate_Person(CurrentSheet, IndexRow)
  End If
  IndexRow = IndexRow + 1
 Loop While Trim(CurrentSheet.Cells(IndexRow, gSrvCol)) <> "<КОНЕЦ>" And IndexRow < 100

End Sub




























