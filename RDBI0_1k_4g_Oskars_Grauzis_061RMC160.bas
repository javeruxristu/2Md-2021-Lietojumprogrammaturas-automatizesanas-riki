Attribute VB_Name = "Module1"
Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Const ClearEmptyDataRowCSS = False
'new requirement @labors ~~ ipzl
'Const PutDefaultsAtCalculationWhenUnableToCalculate = False

'1. Uzdevums
'Kolonna G aprekiniet nostradato stundu skaitu menesi. izmantojot VBA SUB.
Public Sub Calc_Stundas_men()
   Application.StatusBar = "Izpildu 1. uzdevumu..."
   Application.EnableEvents = False
   Config
   Dim Z As Integer, I As Integer
   Z = GetLastUsedRow()
   I = GetFirstDataRow()
   ResetCalcArea ("1")
   For I = I To Z
      Call FillEmptyAndTryFixInvalidDataCells(Range("B" + CStr(I) + ":F" + CStr(I)))
      If (Not IsNumeric(VL(Cells(I, "B"))) Or IsEmpty(VL(Cells(I, "B")))) And (Not IsNumeric(VL(Cells(I, "C"))) Or IsEmpty(VL(Cells(I, "C")))) And (Not IsNumeric(VL(Cells(I, "D"))) Or IsEmpty(VL(Cells(I, "D")))) And (Not IsNumeric(VL(Cells(I, "E"))) Or IsEmpty(VL(Cells(I, "E")))) Then
         If [vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate] Then Cells(I, "G").value = 0
      Else
         Cells(I, "G").value = Val(VL(Cells(I, "B"))) + Val(VL(Cells(I, "C"))) + Val(VL(Cells(I, "D"))) + Val(VL(Cells(I, "E")))
      End If
      If Not ClearEmptyDataRowCSS Then
         'omg es taa ciinijos, kameer dabuuju, lai nonjem nost tos borderus.. Tad atcereejos ka raadiija, ka tacu tieshi uzliek borderus, nevis nonjem...
         Call ApplyBorderToRange(Range("A" + CStr(I) + ":" + "P" + CStr(I)))
      End If
   Next I
   Application.EnableEvents = True
   Application.StatusBar = "1. uzdevums pabeigts!"
   Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

'2. Uzdevums
'Kolonna G ar zalo krasu izceliet tas ?unas, kuras vertiba ir lielaka par 150. Izmantojot teksta krasu.
Public Sub Color_Stundas_men_gt_150()
   Application.StatusBar = "Izpildu 2. uzdevumu..."
   Application.EnableEvents = False
   Config
   Dim Z As Integer, I As Integer
   Z = GetLastUsedRow()
   ResetCalcArea ("2")
   For I = GetFirstDataRow() To Z
      Call FillEmptyAndTryFixInvalidDataCells(Range("B" + CStr(I) + ":F" + CStr(I)))
      If (IsNumeric(VL(Cells(I, "G"))) And Not IsEmpty(VL(Cells(I, "G")))) Then
         If Val(VL(Cells(I, "G"))) > 150 Then
            Cells(I, "G").Font.Color = RGB(0, 255, 0)
         End If
      End If
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("A" + CStr(I) + ":" + "P" + CStr(I)))
      End If

   Next I
   Application.EnableEvents = True
   Application.StatusBar = "2. uzdevums pabeigts!"
   Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

'3. Uzdevums
'Kolonnas H:K noskaidrojiet, vai ir bijusi parstrade katra nedela. Parstrade ir tad, ja darbinieks ir nostradajis vairak par 40 stundam nedela. ?una japarada teksts IR, ja ir bijusi parstrade, un NE, ja parstrade nav bijusi. ?unas, kas satur vardu IR jaizcel ar zalo krasu, ?unas, kas satur vardu Ne, jaizcel ar sarkano krasu.
Public Sub OverWorkChecker()
   Application.StatusBar = "Izpildu 3. uzdevumu..."
   Application.EnableEvents = False
   Config
   Dim value  As String
   Dim Z As Integer, I As Integer
   'for by ref tye mismatch error
   Dim CellSrc As Range, CellDst As Range
   Z = GetLastUsedRow()
   ResetCalcArea ("3")
   For I = GetFirstDataRow() To Z
   Call FillEmptyAndTryFixInvalidDataCells(Range("B" + CStr(I) + ":F" + CStr(I)))
      For j = 0 To 3
         'without set operator it will be autocast to string even with the explict type declaration...
         Set CellSrc = Cells(I, Chr(Asc("B") + j))
         Set CellDst = Cells(I, Chr(Asc("H") + j))
         If (IsNumeric(VL(CellSrc)) And Not IsEmpty(VL(CellSrc))) Then
            value = hasOverwork(VL(CellSrc))
            CellDst.value = value
            If (value = "IR") Then
               CellDst.Interior.Color = RGB(0, 255, 0)
            Else
               CellDst.Interior.Color = RGB(255, 0, 0)
            End If
         Else
            If [vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate] Then
            CellDst.value = "N" + LV("E")
            CellDst.Interior.Color = RGB(255, 0, 0)
            End If
         End If
      Next j
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("A" + CStr(I) + ":" + "P" + CStr(I)))
      End If
   Next I
   Application.EnableEvents = True
   Application.StatusBar = "3. uzdevums pabeigts!"
   Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

'4. Uzdevums
'Kolonnas L:O aprekiniet algu darbiniekam. Ja ir bijusi parstrade, tad darbiniekam samaksa par nostradato stundu ir 150% apmera no maksas par vienu nostradato stundu.
Public Sub PaymentCalculatorWeekly()
   Application.StatusBar = "Izpildu 4. uzdevumu..."
   Application.EnableEvents = False
   Config
   Dim value  As String
   Dim Z As Integer, I As Integer
   Dim CellSrc As Range, CellRate As Range
   Z = GetLastUsedRow()
   ResetCalcArea ("4")
   For I = GetFirstDataRow() To Z
     Call FillEmptyAndTryFixInvalidDataCells(Range("B" + CStr(I) + ":F" + CStr(I)))
      For j = 0 To 3
         Set CellSrc = Cells(I, Chr(Asc("B") + j))
         Set CellDst = Cells(I, Chr(Asc("H") + j))
         Set CellPay = Cells(I, Chr(Asc("L") + j))
         Set CellRate = Cells(I, "F")
         If (IsNumeric(VL(CellSrc)) And Not IsEmpty(VL(CellSrc)) And IsNumeric(VL(CellRate)) And Not IsEmpty(VL(CellRate))) Then
            value = hasOverwork(VL(CellSrc))
            CellDst.value = value
            CellPay.value = calcPaymentWithOverwork(Val(VL(CellSrc)), Val(VL(CellRate)))
            'end numeric source check
         Else
            If [vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate] Then CellPay.value = 0#
         End If
      Next j
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("A" + CStr(I) + ":" + "P" + CStr(I)))
      End If
   Next I
   Application.EnableEvents = True
   Application.StatusBar = "4. uzdevums pabeigts!"
   Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

'5. Uzdevums
'Kolonna P  aprekiniet kopejo algu menesi.  Ja nav iespejams noteikt kopejo algu, tad rezultats ir attelots ar “---“. Krasa sarkana, formatejums Bold.
Public Sub PaymentCalculatorTotal()
   Application.StatusBar = "Izpildu 5. uzdevumu..."
   Application.EnableEvents = False
   Config
   Dim value  As Double
   Dim Z As Integer, I As Integer
   Z = GetLastUsedRow()
   ResetCalcArea ("5")
   For I = GetFirstDataRow() To Z
      Call FillEmptyAndTryFixInvalidDataCells(Range("B" + CStr(I) + ":F" + CStr(I)))
      value = 0#
      Dim hadValue As Boolean
      hadValue = False
      Set dstCell = Cells(I, "P")
      'Here will not transfer to merged value getter for better performance _
      And because this only uses calc area cells which are being unmerged at ResetCalcArea
      For j = 0 To 3
      If (IsNumeric(Cells(I, Chr(Asc("L") + j)).value)) Then
      If (Not IsEmpty(Cells(I, Chr(Asc("L") + j)).value)) Then
      hadValue = True
         value = value + Val(Cells(I, Chr(Asc("L") + j)).value)
         End If
         End If
            Next j

      If (Not hadValue) Then
         dstCell.value = "---"
         dstCell.Font.Color = RGB(255, 0, 0)
         dstCell.Font.Bold = True
      Else
         dstCell.value = value
         dstCell.Font.ColorIndex = 0
         dstCell.Font.Bold = False
      End If
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("A" + CStr(I) + ":" + "P" + CStr(I)))
      End If
   Next I
   Application.EnableEvents = True
   Application.StatusBar = "5. uzdevums pabeigts!"
   Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

Public Sub ResetCalcField()
Application.StatusBar = "T" + LV("i") + "ru apr" + LV("e") + LV("k") + "inu laukumu..."
Application.EnableEvents = False
Dim I As Byte
For I = 0 To 5
ResetCalcArea (CStr(I))
Next
Application.EnableEvents = True
Application.StatusBar = "Apr" + LV("e") + LV("k") + "inu laukums not" + LV("i") + "r" + LV("i") + "ts!"
Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
End Sub

Private Function getOverworkConst() As Integer
   getOverworkConst = 40

End Function

Private Function getOverworkPayMultiplier() As Double
   getOverworkPayMultiplier = 1.5

End Function

Private Function hasOverwork(a) As String
   If IsEmpty(a) Or Not IsNumeric(a) Then
      hasOverwork = "N" + LV("E")
   Else
      If (Val(a) > getOverworkConst()) Then
         hasOverwork = "IR"
      Else
         hasOverwork = "N" + LV("E")
      End If
   End If

End Function

Private Function calcPaymentWithOverwork(hours, rate) As Double
   If (Not IsEmpty(hours) And IsNumeric(hours) And Not IsEmpty(rate) And IsNumeric(rate)) Then
      If (hours > getOverworkConst()) Then
         calcPaymentWithOverwork = (hours - getOverworkConst()) * rate * getOverworkPayMultiplier() + getOverworkConst() * rate
      Else
         calcPaymentWithOverwork = hours * rate
      End If

   Else
      calcPaymentWithOverwork = 0#
   End If
End Function

Function GetFirstDataRow() As Integer
   GetFirstDataRow = 3
End Function

Function GetLastUsedRow(Optional IgnoreLastRow As Boolean = False) As Integer
   Dim x As Integer
   x = ActiveSheet.UsedRange.Rows.Count
   If IgnoreLastRow Then x = x - 1
   For I = x To 1 Step -1
      If (IsEmptyRange(Range("A" + CStr(I) + ":F" + CStr(I)))) Then
         x = I - 1
      Else
         GetLastUsedRow = x
         Exit Function
      End If
   Next I
   If (x = 0) Then x = 1
   GetLastUsedRow = x
End Function

Private Function VL(r As Range) As Variant
If IsObject(r) Then
      If Not r Is Nothing Then
      If r.MergeCells = True Then VL = r.MergeArea.Cells(1, 1).value: Exit Function Else VL = r.value: Exit Function
      End If
      End If
      VL = Empty
End Function

Function ResetCalcArea(Nr As Byte)
   Dim Z As Integer, I As Integer
   Dim j As Integer
   Z = GetLastUsedRow()
   I = GetFirstDataRow()
   Select Case Nr
   Case "1"
   Call UnMergeAllMerged(Range("G" + CStr(I) + ":G" + CStr(Z)))
      With Range("G" + CStr(I) + ":G" + CStr(Z))
         .ClearContents
         .ClearFormats
         .Font.Name = Application.StandardFont
      End With
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("G" + CStr(I) + ":G" + CStr(Z)))
      Else
      For j = I To Z
      If Not IsEmptyRange(Range("A" + CStr(j) + ":F" + CStr(j))) Then Call ApplyBorderToRange(Range("G" + CStr(j) + ":G" + CStr(j)))
      Next j
      End If
      Exit Function
   Case "2"
   Call UnMergeAllMerged(Range("G" + CStr(I) + ":G" + CStr(Z)))
      With Range("G" + CStr(I) + ":G" + CStr(Z))
         .ClearFormats
         .Font.Name = Application.StandardFont
      End With
            If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("G" + CStr(I) + ":G" + CStr(Z)))
      Else
      For j = I To Z
      If Not IsEmptyRange(Range("A" + CStr(j) + ":F" + CStr(j))) Then Call ApplyBorderToRange(Range("G" + CStr(j) + ":G" + CStr(j)))
      Next j
      End If
      Exit Function
   Case "3"
   Call UnMergeAllMerged(Range("H" + CStr(I) + ":K" + CStr(Z)))
      With Range("H" + CStr(I) + ":K" + CStr(Z))
         .ClearContents
         .ClearFormats
         .Font.Name = Application.StandardFont
      End With
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("H" + CStr(I) + ":K" + CStr(Z)))
      Else
      For j = I To Z
      If Not IsEmptyRange(Range("A" + CStr(j) + ":F" + CStr(j))) Then Call ApplyBorderToRange(Range("H" + CStr(j) + ":K" + CStr(j)))
      Next j
      End If
      Exit Function
   Case "4"
   Call UnMergeAllMerged(Range("L" + CStr(I) + ":O" + CStr(Z)))
      With Range("L" + CStr(I) + ":O" + CStr(Z))
         .ClearContents
         .ClearFormats
         .Font.Name = Application.StandardFont
      End With
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("L" + CStr(I) + ":O" + CStr(Z)))
      Else
      For j = I To Z
      If Not IsEmptyRange(Range("A" + CStr(j) + ":F" + CStr(j))) Then Call ApplyBorderToRange(Range("L" + CStr(j) + ":O" + CStr(j)))
      Next j
      End If
      Exit Function
   Case "5"
   Call UnMergeAllMerged(Range("P" + CStr(I) + ":P" + CStr(Z)))
      With Range("P" + CStr(I) + ":P" + CStr(Z))
         .ClearContents
         .ClearFormats
         .Font.Name = Application.StandardFont
      End With
      If Not ClearEmptyDataRowCSS Then
         Call ApplyBorderToRange(Range("P" + CStr(I) + ":P" + CStr(Z)))
      Else
      For j = I To Z
      If Not IsEmptyRange(Range("A" + CStr(j) + ":F" + CStr(j))) Then Call ApplyBorderToRange(Range("P" + CStr(j) + ":P" + CStr(j)))
      Next j
      End If
      Exit Function

   Case Else
   End Select
End Function

Public Sub RemoveInvalidData()
    Application.DisplayStatusBar = True
    Application.StatusBar = False
   Application.StatusBar = "Nodz" + LV("e") + LV("s") + "u neder" + LV("i") + "gos datus..."
   Application.EnableEvents = False
   Config
   Dim Z As Integer, I As Integer, myCell As Range
   Z = GetLastUsedRow()
   I = GetFirstDataRow()
   Call ResetCalcField
   If [vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea] Then
      For I = I To Z
         Set myCell = Cells(I, "A")
         If Not IsEmpty(myCell.value) Then
                Dim oldVal As String, posVal As String
               oldVal = CStr(myCell.value)
            If TryFixLatvianName(oldVal) Then
                  myCell.value = oldVal
               Else
                  myCell.value = ""
               End If
         End If
                  
                  
         For j = 0 To 3
            'without set operator it will be autocast to string even with the explict type declaration...
            Set myCell = Cells(I, Chr(Asc("B") + j))
            If Not IsNumeric(myCell.value) Then myCell.value = ""
            If Not IsEmpty(myCell.value) Then
               If (Val(myCell.value) < 0) Then myCell.value = Abs(Val(myCell.value))
            End If
         Next j
                  
         Set myCell = Cells(I, "F")
         If Not IsNumeric(myCell.value) Then myCell.value = ""
                  
      Next I
   End If
   Application.EnableEvents = True
   Application.StatusBar = "Neder" + LV("i") + "gie dati nodz" + LV("e") + "sti!"
   Application.OnTime Now + TimeValue("00:00:04"), "ClearStatusBar"
End Sub

'Warning: Due to performance, NOT a general usage function, this one has dependancy on data being in work area!
Public Function GetInputAreaRangeCorrespondingToThisRange(r As Range) As Range
If IsObject(r) Then
      If Not r Is Nothing Then
      'first get all the involved rows in the range
      'Never gets called from row(s)/column(s) operations, therefore no performance improvments by first checking whether its full row/col
      'Same thing considering the check whether cell column is outside the work columns
      Dim RowNumbers() As Integer, ptr As Integer, lastNr As Integer, InputRange As Range
      lastNr = 0
            For Each myCell In r
            If Not lastNr = r.Row Then
                     ptr = getNextIndexForIntegerArray(RowNumbers)
                     ReDim Preserve RowNumbers(ptr)
                     RowNumbers(ptr) = r.Row
                     lastNr = r.Row
                  End If
            Next

            'cant simply use ptr, because it will be 0 when array has one element and also when there are no elements
            Dim ArrayDotLength As Integer
            ArrayDotLength = getNextIndexForIntegerArray(RowNumbers)
            If ArrayDotLength > 0 Then
        'second, if any rows then generate the input ranges and return
        'Set InputRange = Union(InputRange,[C12,K22,N10,G11])
        For I = 0 To ArrayDotLength - 1
        If InputRange Is Nothing Then Set InputRange = Range("A" + CStr(RowNumbers(I)) + ":F" + CStr(RowNumbers(I)))
        Set InputRange = Union(InputRange, Range("A" + CStr(RowNumbers(I)) + ":F" + CStr(RowNumbers(I))))
        Next
        Set GetInputAreaRangeCorrespondingToThisRange = InputRange
        Exit Function
        End If
      End If
End If

End Function

Public Function TryFixLatvianName(nm As String) As Boolean
   Dim oldVal As String, posVal As String
   Dim str() As String
   oldVal = Trim(nm)
   If Not IsLVNameValid(oldVal) Then
      str = SplitRe(oldVal, "[\s]")
      If (UBound(str) - LBound(str)) < 1 Then
         posVal = Trim(UCase(Mid(oldVal, 1, 1)) + Mid(oldVal, 2, Len(oldVal) - 1))
         If IsLVNameValid(posVal) Then
            TryFixLatvianName = True
            nm = posVal
            Exit Function
         End If
      Else
         Dim LenArr As Integer, posArr() As String, ptr As Integer
         LenArr = UBound(str) - LBound(str) '+ 1
         For I = 0 To LenArr
            If Not IsEmpty(str(I)) Then
               ptr = getNextIndexForArray(posArr)
               ReDim Preserve posArr(ptr)
               posArr(ptr) = UCase(Mid(str(I), 1, 1)) + Mid(str(I), 2, Len(str(I)) - 1)
            End If
         Next I
         posVal = Join(posArr, " ")
         If IsLVNameValid(posVal) Then
            TryFixLatvianName = True
            nm = posVal
            Exit Function
         End If
      End If
   Else
      TryFixLatvianName = True
      nm = oldVal
      Exit Function
   End If
   TryFixLatvianName = False
End Function

Public Function FillEmptyAndTryFixInvalidDataCells(r As Range, Optional HasBeenCalledAtInput = False)
   'Mid(myCell.Address, 2, 1) = "F" (well for 1 letter columns only)
   'Mid(myCell.Address, 4, 2) <> "8" (for up to 2 row number rows)
   Dim ShouldDoAutoFill As Boolean
   ShouldDoAutoFill = ([vba_setting_AutoFillEmptyDataAreasAtInput] And HasBeenCalledAtInput) _
            Or ([vba_setting_AutoFillEmptyDataAreasAtCalc] And Not HasBeenCalledAtInput)
   If IsObject(r) Then
      If Not r Is Nothing Then
         If ShouldDoAutoFill Then
            Dim myCell As Range
            Dim c As String
            Dim hadEventsEnabled As Boolean
            hadEventsEnabled = Application.EnableEvents
            If hadEventsEnabled Then Application.EnableEvents = False
            For Each myCell In r
               'NB! should check just as in applycss method by actual splitting etc.  - for 2+ letter columns _
               as here A, AA, AB, AC, AAD will all pass
               c = Mid(myCell.Address, 2, 1)
               If (c = "A") Then
                  If Not IsEmpty(myCell.value) And [vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea] Then
                  Dim oldVal As String
                  oldVal = CStr(myCell.value)
                     If Not TryFixLatvianName(oldVal) Then
                           myCell.value = ""
                           Else
                           myCell.value = oldVal
                     End If
                  End If
               ElseIf Asc(c) > Asc("A") And Asc(c) <= Asc("E") Then
                  If Not IsNumeric(myCell.value) Or IsEmpty(myCell.value) Then
                     myCell.value = 0
                  Else
                     If Val(myCell.value) < 0 And [vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea] Then myCell.value = Abs(Val(myCell.value))
                  End If
               ElseIf c = "F" Then
                  If [vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea] Then
                     If Not IsNumeric(myCell.value) Or IsEmpty(myCell.value) Then myCell.value = 0#
                  Else
                     If IsEmpty(myCell.value) Then myCell.value = 0#
                  End If
               End If
            Next myCell
            'reduce the if check here to simple assignment back to the original value
            Application.EnableEvents = hadEventsEnabled
         End If
      End If
   End If
End Function



' ByVal r      As Range
Public Sub ApplyCss(ByVal r As Range)
   ' Dim r    As Range
   ' Set r = Range("A1:P9")
   If IsObject(r) Then
      If Not r Is Nothing Then
         If (r.Column <= Asc("P") - Asc("A") + 1) Then
            Dim endRow, endCol, rRange As Range
            If (IndexOf(r.Address, ":") > -1) Then
               Dim rowStart, rowEnd, ArrParts() As String
               ArrParts = Split(r.Address, ":")
               rowStart = Mid(ArrParts(0), 2, Len(ArrParts(0)) - 1)
               rowEnd = Mid(ArrParts(1), 2, Len(ArrParts(1)) - 1)
               If IsNumeric(rowStart) And IsNumeric(rowEnd) Then
                  'row insert or undo delete operation
                  Exit Sub
                  'further proccessing will result into unable undo!!!
                  'endCol = Asc("P") - Asc("A") + 1
                  'endRow = rowEnd
               Else
                  endCol = Asc(Split(ArrParts(1), "$")(1)) - Asc("A") + 1

                  endRow = Split(ArrParts(1), "$")(2)
               End If
            Else
               endRow = r.Row
               endCol = r.Column
            End If

            'ignore values past the work area
            If endCol > Asc("F") - Asc("A") + 1 Then
               endCol = Asc("F") - Asc("A") + 1
            End If

            Dim I As Integer
            For I = 0 To endRow - r.Row
               Dim vRange As Range, inRange As Range
               Set vRange = Range(Chr(Asc("A") - 1 + r.Column) + CStr(r.Row + I) + ":" + Chr(Asc("A") - 1 + endCol) + CStr(r.Row + I))
               If [vba_setting_AutoFillEmptyDataAreasAtInput] Then Set vRange = Range(Chr(Asc("A") - 1 + r.Column) + CStr(r.Row + I) + ":F" + CStr(r.Row + I))
               If Not IsEmptyRange(vRange) Then
               If Not ClearEmptyDataRowCSS And I = 0 Then Set rRange = Range("A" + CStr(GetLastUsedRow(True)) + ":" + "P" + CStr(r.Row + I)) _
               Else Set rRange = Range("A" + CStr(r.Row + I) + ":" + "P" + CStr(r.Row + I))
                  Call ApplyBorderToRange(rRange)
               Else
                  If Not ClearEmptyDataRowCSS Then
                  Else
                     'cases when some cols only changed, have to check all the rest of the input cols
                     Set inRange = Range("A" + CStr(r.Row + I) + ":" + "F" + CStr(r.Row + I))
                     If IsEmptyRange(inRange) Then
                        Call RemoveCss(inRange)
                     End If
                  End If
               End If
            Next I
         End If
      End If
   End If
End Sub

Private Function ApplyBorderToRange(r As Range)
   If IsObject(r) Then
      If Not r Is Nothing Then
         With r.Borders
            .LineStyle = 1
            .Color = 0
            .Weight = 2
         End With
      End If
   End If

End Function

Public Sub RemoveCss(ByVal r As Range)
   ' Dim r    As Range
   ' Set r = Range("A1:P9")

   If IsObject(r) Then
      If Not r Is Nothing Then
         If (r.Column <= Asc("P") - Asc("A") + 1) Then
            'If (Not IsEmptyRange(r)) Then
            Dim endRow
            Dim ArrParts() As String
            If (IndexOf(r.Address, ":") > -1) Then
               ArrParts = Split(r.Address, ":")
               Dim rowStart As String, rowEnd As String
               rowStart = Mid(ArrParts(0), 2, Len(ArrParts(0)) - 1)
               rowEnd = Mid(ArrParts(1), 2, Len(ArrParts(1)) - 1)
               If rowStart = rowEnd And IsNumeric(rowStart) Then
                  endRow = r.Row
                  Exit Sub
                  'further proccessing will result into unable undo!!!
               Else
                  endRow = CInt(Split(ArrParts(1), "$")(2))
               End If
            Else
               endRow = r.Row

            End If
            ' MsgBox (Range("A2").Borders)
            Dim prevRowRange As Range, nextRowRange As Range, rRange As Range
            Dim prevRowEmpty As Boolean, nextRowEmpty As Boolean, currRowEmpty As Boolean
            Dim I As Integer
            For I = 0 To endRow - r.Row
               Set rRange = Range("A" + CStr(r.Row + I) + ":" + "F" + CStr(r.Row + I))
               Set nextRowRange = Range("A" + CStr(r.Row + I + 1) + ":" + "F" + CStr(r.Row + I + 1))

               If (r.Row + I - 1 > 0) Then
                  If (I = 0) Then
                     ' sobriid bezjeedziigi, bet kad buus sadaliitie gabali, noderees!
                     Set prevRowRange = Range("A" + CStr(r.Row + I - 1) + ":" + "F" + CStr(r.Row + I - 1))
                     prevRowEmpty = IsEmptyRange(prevRowRange)
                     currRowEmpty = IsEmptyRange(rRange)
                  Else
                     prevRowEmpty = currRowEmpty
                     currRowEmpty = nextRowEmpty
                  End If
               Else
                  prevRowEmpty = True
               End If

               Set nextRowRange = Range("A" + CStr(r.Row + I + 1) + ":" + "F" + CStr(r.Row + I + 1))
               nextRowEmpty = IsEmptyRange(nextRowRange)
               Set rRange = Range("A" + CStr(r.Row + I) + ":" + "P" + CStr(r.Row + I))
               Call RemoveBorderToRange(rRange, prevRowEmpty, currRowEmpty, nextRowEmpty)

            Next I
         End If
      End If
   End If
End Sub

' ByVal r      As Range
Private Function RemoveBorderToRange(ByVal r As Range, ClearTop As Boolean, ClearLeftRight As Boolean, ClearBottom As Boolean)
   If IsObject(r) Then
      If Not r Is Nothing Then
         'r.Select
         'unable to set xlNone or dot (-4247) if other values like weight and colorindex are being set!
         If ClearTop Then
            r.Borders(xlEdgeTop).LineStyle = xlNone
         End If
         If ClearLeftRight Then
            r.Borders(xlInsideVertical).LineStyle = xlNone
            r.Borders(xlEdgeRight).LineStyle = xlNone
            r.Borders(xlEdgeLeft).LineStyle = xlNone
         End If
         If ClearBottom Then
            r.Borders(xlBottom).LineStyle = xlNone
         End If
      End If
   End If

End Function

Public Sub AdjustValidation(ByVal r As Range)
   If IsObject(r) Then
      If Not r Is Nothing Then
         If (r.Column <= Asc("P") - Asc("A") + 1) Then
            Dim endRow, rRange As Range
            If (IndexOf(r.Address, ":") > -1) Then
               Dim ArrParts() As String
               ArrParts = Split(r.Address, ":")
               rowStart = Mid(ArrParts(0), 2, Len(ArrParts(0)) - 1)
               rowEnd = Mid(ArrParts(1), 2, Len(ArrParts(1)) - 1)
               If IsNumeric(rowStart) And IsNumeric(rowEnd) Then
                  'row insert or undo delete operation
                  Exit Sub
                  'further proccessing will result into unable undo!!!
                  'endRow = rowEnd
               Else
               'This sub will never get called on column operation, therefore this means that this is cell operation not row
                  endRow = Split(ArrParts(1), "$")(2)
               End If
            Else
               endRow = r.Row
            End If
            Set rRange = Range("A" + CStr(r.Row) + ":" + "F" + CStr(endRow))
            Dim oldValue
            Dim myCell As Range
            Dim isEmptyAllFields As Boolean
            If ClearEmptyDataRowCSS Then isEmptyAllFields = IsEmptyRange(rRange)

            For Each myCell In rRange
            Dim HasValidation As Boolean
            HasValidation = HasCellValidation(myCell)
               If Not HasValidation And Not ClearEmptyDataRowCSS Then
                  'omg
                  '7 _
                  You cannot use a UDF directly in data validation. You can however use it via a named formula. _
                  Select A1, then in Name Manager define a name called IsAlphaNum whose refersto is: _
                  =alphanumeric(A1)
                  ' https://stackoverflow.com/a/36470614 _
                  FAILS if reference cell is put in a veryhidden sheet!

                  'so constant... MinDecimal(0) in cell shows -79228162514264300000000000000
                  ' myCell.Select
                  If myCell.Column > Asc("E") - Asc("A") + 1 Then

                     With myCell.Validation
                        .Add Type:=xlValidateDecimal, _
                             AlertStyle:=xlValidAlertStop, _
                             Operator:=xlGreaterEqual, Formula1:="-79228162514264300000000000000"
                        .InputTitle = "Skait" + LV("l") + "i"
                        .ErrorTitle = "Skait" + LV("l") + "i"
                        .ShowInput = True
                        .InputMessage = "Ievadiet skaitli"
                        .ErrorMessage = "Jums j" + LV("a") + "ievada skaitlis"
                        .IgnoreBlank = True
                     End With
                  Else
                     If myCell.Column = 1 Then
                        Dim addr As String
                        addr = Replace(myCell.Address, "$", "")
                                                            With myCell.Validation
                                        .Add Type:=xlValidateCustom, _
                                             AlertStyle:=xlValidAlertStop, _
                                             Formula1:="=OR(" + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""A""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""B""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""C""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""D""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""E""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""F""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""G""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""H""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""I""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""J""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""K""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""L""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""M""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""N""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""O""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""P""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""Q""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""R""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""S""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""T""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""U""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""V""))," + _
                                             "(EXACT(MID(TRIM(" + addr + "),1,1), ""W""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""X""))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), ""Y""))," + "(EXACT(MID(TRIM(" + addr + "),1,1), ""Z""))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("A") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("C") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("E") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("G") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("I") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("K") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("L") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("N") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("O") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("R") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("S") + """))," + "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("U") + """))," + _
                           "(EXACT(MID(TRIM(" + addr + "),1,1), """ + LV("Z") + """))" + _
                           ")"
                           .InputTitle = "V" + LV("a") + "rdam j" + LV("a") + "s" + LV("a") + "kas ar lielo burtu"
                           .ErrorTitle = "V" + LV("a") + "rdam j" + LV("a") + "s" + LV("a") + "kas ar lielo burtu"
                           .ShowInput = True
                           .InputMessage = "Ievadiet darbinieka v" + LV("a") + "rdu (s)"
                           .ErrorMessage = "Pirmajam ievad" + LV("i") + "tajam burtam j" + LV("a") + "b" + LV("u") + "t lielajam" + _
                                                                                                                   vbCrLf + "piem" + LV("e") + "ram: " + vbCrLf + _
                                                                                                                   "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + _
                                                                                                                   LV("A") + LV("C") + _
                                                                                                                   LV("E") + LV("G") + _
                                                                                                                   LV("I") + LV("K") + _
                                                                                                                   LV("L") + LV("N") + _
                                                                                                                   LV("O") + LV("R") + _
                                                                                                                   LV("S") + LV("U") + _
                                                                                                                   LV("Z")
                           .IgnoreBlank = True
                        End With
                     Else
                        With myCell.Validation
                           .Add Type:=xlValidateDecimal, _
                                AlertStyle:=xlValidAlertStop, _
                                Operator:=xlGreaterEqual, Formula1:="0"
                           .InputTitle = "Nenegat" + LV("i") + "vi skait" + LV("l") + "i"
                           .ErrorTitle = "Nenegat" + LV("i") + "vi skait" + LV("l") + "i"
                           .ShowInput = True
                           .InputMessage = "Ievadiet stundu skaitu 0 vai vair" + LV("a") + "k"
                           .ErrorMessage = "Jums j" + LV("a") + "ievada skaitlis 0 vai vair" + LV("a") + "k"
                           .IgnoreBlank = True
                        End With
                     End If
                  End If
               Else
                  If isEmptyAllFields Then
                     'will remove the validations also if option turned on. So that data is easier to look at
                     If ClearEmptyDataRowCSS And HasValidation Then
                        With myCell.Validation
                           .Delete
                        End With
                     End If
                  End If
               End If
            Next myCell
         End If
      End If
   End If
End Sub

Public Sub RemoveValidationAtRowInsert(ByVal r As Range)
   'still the validation is entirely different matter than borders. _
   but in this case there are no more empty title rows left, so therefore _
   no point in processing the validations
   If Not ClearEmptyDataRowCSS Then
   Else
      If IsObject(r) Then
         If Not r Is Nothing Then
            If (r.Column <= Asc("P") - Asc("A") + 1) Then

               If (IndexOf(r.Address, ":") > -1) Then
                  Dim ArrParts() As String
                  ArrParts = Split(r.Address, ":")
                  If (UBound(ArrParts) - LBound(ArrParts) + 1 = 2) Then
                     Dim rowStart As String, rowEnd As String
                     rowStart = Mid(ArrParts(0), 2, Len(ArrParts(0)) - 1)
                     rowEnd = Mid(ArrParts(1), 2, Len(ArrParts(1)) - 1)
                     If rowStart = rowEnd And IsNumeric(rowStart) Then
                        'A Row is being inserted!
                        Dim rRange As Range
                        Dim myCell As Range
                        Set rRange = Range("A" + CStr(r.Row) + ":" + "F" + CStr(r.Row))
                        For Each myCell In rRange
                           If (HasCellValidation(myCell)) Then
                              With myCell.Validation
                                 .Delete
                              End With
                           End If
                        Next myCell
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Public Function ValidateLatvianNames(r As Range) As Boolean
   If IsObject(r) Then
      If Not r Is Nothing Then
         If r.Column = 1 Then
            Dim invalidCells() As String, ptr As Byte, hasNoValidationObjections As Variant
            Dim myCell As Range
            For Each myCell In r
              If myCell.Column = 1 And Not IsEmpty(myCell.value) Then
               If HasCellValidation(myCell) Then hasNoValidationObjections = myCell.Validation.value _
               Else hasNoValidationObjections = True
               If hasNoValidationObjections Then
                If Not IsLVNameValid(Trim(CStr(myCell.value))) Then
                     ptr = getNextIndexForArray(invalidCells)
                     ReDim Preserve invalidCells(ptr)
                     invalidCells(ptr) = LV("S") + LV("u") + "na " + Replace(myCell.Address, "$", "")
                  End If
                End If
               End If
            Next

            'output invalids
            If getNextIndexForArray(invalidCells) > 0 Then
               Dim msgString As String
               msgString = "A kolon" + LV("a") + " atrasti nepareizi dati : " + vbCrLf + _
                                               "L" + LV("u") + "dzu ievadiet v" + LV("a") + "rdu(s), kas s" + LV("a") + "kas ar Lielo burtu un turpin" + LV("a") + "s ar mazajiem!" + vbCrLf
               For I = 0 To UBound(invalidCells) - LBound(invalidCells)
                  msgString = msgString + CStr(invalidCells(I))
                  If UBound(invalidCells) - LBound(invalidCells) + 1 - I > 1 Then
                     If UBound(invalidCells) - LBound(invalidCells) + 1 - I > 2 Then
                        msgString = msgString + "    " + CStr(invalidCells(I + 1)) + "    " + CStr(invalidCells(I + 2)) + vbCrLf
                        I = I + 2
                     Else
                        msgString = msgString + "    " + CStr(invalidCells(I + 1)) + vbCrLf
                        I = I + 1
                     End If
                  End If
               Next I
               MsgBoxW (msgString)
            Else
               ValidateLatvianNames = True
               Exit Function
            End If

         End If
      End If
   End If
   ValidateLatvianNames = False
End Function

Public Function IsLVNameValid(regexStr As String) As Boolean
   Dim regex  As Object
   Set regex = CreateObject("VBScript.RegExp")
   'Dim matchesVal As Object
   Dim lvNameRegexPattern As String
   lvNameRegexPattern = "^([A-Z" + LV("A") + LV("C") + LV("E") + LV("G") + LV("I") + LV("K") + LV("L") + LV("N") + LV("O") + LV("R") + LV("S") + LV("U") + LV("Z") + "][a-zA-Z" + LV("a") + LV("A") + LV("c") + LV("C") + LV("e") + LV("E") + LV("g") + LV("G") + LV("i") + LV("I") + LV("k") + LV("K") + LV("l") + LV("L") + LV("n") + LV("N") + LV("o") + LV("O") + LV("r") + LV("R") + LV("s") + LV("S") + LV("u") + LV("U") + LV("z") + LV("Z") + "]+)" + _
                                                                                                                                                                                                                                                                                                                                                                                                                                                    "(\s[A-Z" + LV("A") + LV("C") + LV("E") + LV("G") + LV("I") + LV("K") + LV("L") + LV("N") + LV("O") + LV("R") + LV("S") + LV("U") + LV("Z") + "][a-zA-Z" + LV("a") + LV("A") + LV("c") + LV("C") + LV("e") + LV("E") + LV("g") + LV("G") + LV("i") + LV("I") + LV("k") + LV("K") + LV("l") + LV("L") + LV("n") + LV("N") + LV("o") + LV("O") + LV("r") + LV("R") + LV("s") + LV("S") + LV("u") + LV("U") + LV("z") + LV("Z") + "]+)*$"

   'lvNameRegexPattern = "([a-zA-Z" + LV("a") + LV("A") + LV("c") + LV("C") + LV("e") + LV("E") + LV("g") + LV("G") + LV("i") + LV("I") + LV("k") + LV("K") + LV("l") + LV("L") + LV("n") + LV("N") + LV("o") + LV("O") + LV("r") + LV("R") + LV("s") + LV("S") + LV("u") + LV("U") + LV("z") + LV("Z") + "]+)" + _
   "(\s[a-zA-Z" + LV("a") + LV("A") + LV("c") + LV("C") + LV("e") + LV("E") + LV("g") + LV("G") + LV("i") + LV("I") + LV("k") + LV("K") + LV("l") + LV("L") + LV("n") + LV("N") + LV("o") + LV("O") + LV("r") + LV("R") + LV("s") + LV("S") + LV("u") + LV("U") + LV("z") + LV("Z") + "]+)*"
   With regex
      .Pattern = lvNameRegexPattern
      .Global = False
      .IgnoreCase = False
      .MultiLine = False
   End With
   'Set matchesVal = regex.Execute(regexStr)
   'For Each Item In matchesVal
   '    Debug.Print Item.value
   'Next
   IsLVNameValid = regex.test(regexStr)
End Function

Public Function LV(c As String) As String
   Select Case c
   Case "a"
      LV = ChrW(257)
      Exit Function
   Case "A"
      LV = ChrW(256)
      Exit Function
   Case "c"
      LV = ChrW(269)
      Exit Function
   Case "C"
      LV = ChrW(268)
      Exit Function
   Case "e"
      LV = ChrW(275)
      Exit Function
   Case "E"
      LV = ChrW(274)
      Exit Function
   Case "g"
      LV = ChrW(289)
      Exit Function
   Case "G"
      LV = ChrW(286)
      Exit Function
   Case "i"
      LV = ChrW(299)
      Exit Function
   Case "I"
      LV = ChrW(298)
      Exit Function
   Case "k"
      LV = ChrW(311)
      Exit Function
   Case "K"
      LV = ChrW(310)
      Exit Function
   Case "l"
      LV = ChrW(316)
      Exit Function
   Case "L"
      LV = ChrW(315)
      Exit Function
   Case "n"
      LV = ChrW(326)
      Exit Function
   Case "N"
      LV = ChrW(325)
      Exit Function
   Case "o"
      LV = ChrW(333)
      Exit Function
   Case "O"
      LV = ChrW(332)
      Exit Function
   Case "r"
      LV = ChrW(343)
      Exit Function
   Case "R"
      LV = ChrW(342)
      Exit Function
   Case "s"
      LV = ChrW(353)
      Exit Function
   Case "S"
      LV = ChrW(352)
      Exit Function
   Case "u"
      LV = ChrW(363)
      Exit Function
   Case "U"
      LV = ChrW(362)
      Exit Function
   Case "z"
      LV = ChrW(382)
      Exit Function
   Case "Z"
      LV = ChrW(381)
      Exit Function

   Case Else
   End Select
   LV = c
End Function

Public Function IndexOf(haystack As String, needle As String)
   Dim res
   res = InStr(haystack, needle)
   If Not IsNumeric(res) Then
      IndexOf = -1
   Else
      If res = 0 Then
         IndexOf = -1
      Else
         IndexOf = res
      End If
   End If
End Function

'WTF can not simply state as Variant for all the array types, have to write for each type seperately?????
Function getNextIndexForArray(a() As String) As Byte
   If (Not a) = -1 Then
      ' Array has NOT been initialized
      getNextIndexForArray = 0
   Else
      getNextIndexForArray = UBound(a) + 1
   End If
End Function

Function getNextIndexForVariantArray(a() As Variant) As Byte
   If (Not a) = -1 Then
      ' Array has NOT been initialized
      getNextIndexForVariantArray = 0
   Else
      getNextIndexForVariantArray = UBound(a) + 1
   End If
End Function

Function getNextIndexForIntegerArray(a() As Integer) As Byte
   If (Not a) = -1 Then
      ' Array has NOT been initialized
      getNextIndexForIntegerArray = 0
   Else
      getNextIndexForIntegerArray = UBound(a) + 1
   End If
End Function

'Why AskConfiguration is not in the assignable macros list? It has no parameter by default - it is like a sub with one "overload"...
Public Sub ReConfigure()
Call AskConfiguration
End Sub

Public Sub AskConfiguration(Optional IsOnOpen As Boolean = False)
   Dim TryConvertOrRemoveInvalidEntriesFromDataArea As Boolean, AutoFillEmptyDataAreasAtInput As Boolean, _
       AutoFillEmptyDataAreasAtCalc As Boolean, PutDefaultsAtCalculationWhenUnableToCalculate As Boolean
   Dim rsp As String, promptString As String
   promptString = "Labdien! l" + LV("u") + "dzu atbildiet uz "
   If IsOnOpen Then promptString = promptString + CStr(3) Else promptString = promptString + CStr(4)
   promptString = promptString + " jaut" + LV("a") + "jumiem, lai nokonfigur" + LV("e") + "tu " + LV("s") + _
                                    "o sessiju. " + vbCrLf + vbCrLf + "1. Vai "
   If IsOnOpen Then promptString = promptString + "tagad izdz" + LV("e") + "st k" + LV("l") + LV("u") + "dainos datus no ievades datu tabulas un ar" + LV("i")
   promptString = promptString + " saskaroties ar nekorektiem datiem " + LV("A") + "rpus valid" + LV("a") + "cijas" + " (piem" + LV("e") + "ram, stundu skaits negat" + LV("i") + "vs) tos m" + LV("e") + LV("g") + "in" + LV("a") + "t reduc" + LV("e") + "t uz noklus" + LV("e") + "jumu ( piem" + LV("e") + "ram, uz negativ" + LV("a") + "s v" + LV("e") + "rt" + LV("i") + "bas moduli)?"

   rsp = MsgBoxW(promptString, vbYesNo)
   If rsp = vbYes Then
      TryConvertOrRemoveInvalidEntriesFromDataArea = True
   Else
      TryConvertOrRemoveInvalidEntriesFromDataArea = False
   End If
   rsp = MsgBoxW("2. Vai autom" + LV("a") + "tiski aizpild" + LV("i") + "t tuk" + LV("s") + "os datu laukus kad dati tiek ievad" + LV("i") + "ti?", vbYesNo)
   If rsp = vbYes Then
      AutoFillEmptyDataAreasAtInput = True
   Else
      AutoFillEmptyDataAreasAtInput = False
   End If
   rsp = MsgBoxW("3. Vai autom" + LV("a") + "tiski aizpild" + LV("i") + "t tuk" + LV("s") + "os datu laukus kad dati tiek apr" + LV("e") + LV("k") + "in" + LV("a") + "ti?", vbYesNo)
   If rsp = vbYes Then
      AutoFillEmptyDataAreasAtCalc = True
   Else
      AutoFillEmptyDataAreasAtCalc = False
   End If
   
   If (IsOnOpen) Then
   Call Config(TryConvertOrRemoveInvalidEntriesFromDataArea, AutoFillEmptyDataAreasAtInput, AutoFillEmptyDataAreasAtCalc)
   If IsOnOpen And TryConvertOrRemoveInvalidEntriesFromDataArea Then
   Dim Target As Range
   Dim Z As Integer, I As Integer, myCell As Range
   Z = GetLastUsedRow()
   I = GetFirstDataRow()
   Call RemoveInvalidData
   Set Target = Range("A" + CStr(I) + ":F" + CStr(Z))
   Call AdjustValidation(Target)
   End If
   Else
      rsp = MsgBoxW("4. Vai gad" + LV("i") + "jumos, kad nav iesp" + LV("e") + "jams apr" + LV("e") + LV("k") + "in" + LV("a") + "t rezult" + LV("a") + "tu, aizpild" + LV("i") + "t tuk" + LV("s") + "os rezult" + LV("a") + "tus ar noklus" + LV("e") + "juma v" + LV("e") + "rt" + LV("i") + "b" + LV("a") + "m?", vbYesNo)
   If rsp = vbYes Then
      PutDefaultsAtCalculationWhenUnableToCalculate = True
   Else
      PutDefaultsAtCalculationWhenUnableToCalculate = False
   End If
   Call Config(TryConvertOrRemoveInvalidEntriesFromDataArea, AutoFillEmptyDataAreasAtInput, AutoFillEmptyDataAreasAtCalc, PutDefaultsAtCalculationWhenUnableToCalculate)
   End If
End Sub

'should do overload here, because if settings are present, then need to know whether the parameters should
'override them ( they were specified in this function call)
'or whether the existing settings should override the parameters
'returns true if already configured, false if had to create the config
Public Function Config(Optional TryConvertOrRemoveInvalidEntriesFromDataArea As Variant, _
                       Optional AutoFillEmptyDataAreasAtInput As Variant, _
                       Optional AutoFillEmptyDataAreasAtCalc As Variant, _
                       Optional PutDefaultsAtCalculationWhenUnableToCalculate As Variant) _
                       As Boolean
   Dim hadError As Byte, wasSpecified As Byte
   hadError = 0
   wasSpecified = 0

   If Not IsMissing(TryConvertOrRemoveInvalidEntriesFromDataArea) Then wasSpecified = wasSpecified Or 1
   If Not IsMissing(AutoFillEmptyDataAreasAtInput) Then wasSpecified = wasSpecified Or 2
   If Not IsMissing(AutoFillEmptyDataAreasAtCalc) Then wasSpecified = wasSpecified Or 4
   If Not IsMissing(PutDefaultsAtCalculationWhenUnableToCalculate) Then wasSpecified = wasSpecified Or 8

   'defaults
   If Not ((wasSpecified And 1) = 1) Then TryConvertOrRemoveInvalidEntriesFromDataArea = False
   If Not ((wasSpecified And 2) = 2) Then AutoFillEmptyDataAreasAtInput = False
   If Not ((wasSpecified And 4) = 4) Then AutoFillEmptyDataAreasAtCalc = False
   If Not ((wasSpecified And 8) = 8) Then PutDefaultsAtCalculationWhenUnableToCalculate = True


   If IsError([vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea]) Then
      ThisWorkbook.Names.Add "vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea", TryConvertOrRemoveInvalidEntriesFromDataArea
      hadError = hadError Or 1
   End If
   If IsError([vba_setting_AutoFillEmptyDataAreasAtInput]) Then
      ThisWorkbook.Names.Add "vba_setting_AutoFillEmptyDataAreasAtInput", AutoFillEmptyDataAreasAtInput
      hadError = hadError Or 2
   End If

   If IsError([vba_setting_AutoFillEmptyDataAreasAtCalc]) Then
      ThisWorkbook.Names.Add "vba_setting_AutoFillEmptyDataAreasAtCalc", AutoFillEmptyDataAreasAtCalc
      hadError = hadError Or 4
   End If
   
      If IsError([vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate]) Then
      ThisWorkbook.Names.Add "vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate", PutDefaultsAtCalculationWhenUnableToCalculate
      hadError = hadError Or 8
   End If


   If hadError = 0 And wasSpecified = 0 Then Config = True: Exit Function

   If wasSpecified > 0 Then
      'why evaluate read works, but write not??
      'tBool = [vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea]

      'now I am forced to write in long way such as:
      'ThisWorkbook.Names.Item("vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea").value = TryConvertOrRemoveInvalidEntriesFromDataArea

      'instead of more convient :
      'Why these throw object required error??
      '[vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea] = TryConvertOrRemoveInvalidEntriesFromDataArea
      '["vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea"] = TryConvertOrRemoveInvalidEntriesFromDataArea
      'Evaluate("vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea") = TryConvertOrRemoveInvalidEntriesFromDataArea
      If wasSpecified And 1 = 1 Then ThisWorkbook.Names.Item("vba_setting_TryConvertOrRemoveInvalidEntriesFromDataArea").value = TryConvertOrRemoveInvalidEntriesFromDataArea
      If wasSpecified And 2 = 2 Then ThisWorkbook.Names.Item("vba_setting_AutoFillEmptyDataAreasAtInput").value = AutoFillEmptyDataAreasAtInput
      If wasSpecified And 4 = 4 Then ThisWorkbook.Names.Item("vba_setting_AutoFillEmptyDataAreasAtCalc").value = AutoFillEmptyDataAreasAtCalc
      If wasSpecified And 8 = 8 Then ThisWorkbook.Names.Item("vba_setting_PutDefaultsAtCalculationWhenUnableToCalculate").value = PutDefaultsAtCalculationWhenUnableToCalculate
      Config = False
      Application.DisplayStatusBar = True
      'sometimes excel is using the statusbar and results in error...
      On Error Resume Next
      If hadError > 0 Then Application.StatusBar = "Please leave all the names that start with vba_setting_ intact. They are required for normal program operation!": Application.OnTime Now + TimeValue("00:00:04"), "ClearStatusBar"
      Exit Function
   Else

   End If
   Config = False
   Application.DisplayStatusBar = True
   Application.StatusBar = "Please leave all the names that start with vba_setting_ intact. They are required for normal program operation!"
   Application.OnTime Now + TimeValue("00:00:04"), "ClearStatusBar"
End Function

' Thank you : Tushar Mehta
' for this function : http://www.tushar-mehta.com/publish_train/xl_vba_cases/1003%20MinMaxVals.shtml
Function MaxDecimal(ByVal NbrDecimals As Byte) As Variant

   'there is a bug in the code that handles decimal division so that _
   even when the resulting value fits into a decimal data type, the _
   software rounds the last digit. Hence, the correction below.

   MaxDecimal = CDec("79,228,162,514,264,337,593,543,950,335")

   If NbrDecimals > 0 Then

      MaxDecimal = MaxDecimal / CDec(10) ^ CDec(NbrDecimals)

      Dim Correction

      Correction = CDec(0.5) / CDec(10) ^ (CDec(NbrDecimals) - 1)

      MaxDecimal = MaxDecimal - Correction

   End If

End Function

' Thank you : Tushar Mehta
' for this function : http://www.tushar-mehta.com/publish_train/xl_vba_cases/1003%20MinMaxVals.shtml
Function MinDecimal(ByVal NbrDecimals As Byte) As Variant

   MinDecimal = -MaxDecimal(NbrDecimals)

End Function

'nesaturees vbNullChar taapeec varu drosi izmantot so!
Public Function SplitRe(Text As String, Pattern As String, Optional IgnoreCase As Boolean) As String()
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.MultiLine = True
    End If

    re.IgnoreCase = IgnoreCase
    re.Pattern = Pattern
    SplitRe = Strings.Split(re.Replace(Text, ChrW(-1)), ChrW(-1))
End Function


Public Function IsFromTheInputRange(ByVal r As Range) As Boolean
   ' Dim r    As Range
   ' Set r = Range("A1:P9")
   If IsObject(r) Then
      If Not r Is Nothing Then

         Dim endCol As String, endRow As String, rRange As Range
         'This will only get called NOT on full row/col operations, therefore safe to skip that case checking here
         If (IndexOf(r.Address, ":") > -1) Then
            Dim ArrParts() As String
            ArrParts = Split(r.Address, ":")
            'splitting on $ $a$1111 results in an emty 0th result
            endRow = Split(ArrParts(1), "$")(2)
            endCol = Split(ArrParts(1), "$")(1)
         Else
            endCol = r.Column
            endRow = r.Row
         End If
         'If user has tried to change anything at calculation field
         If ((r.Column > Asc("G") - Asc("A") And (r.Column <= Asc("P") - Asc("G") + Asc("G") - Asc("A") + 1)) Or (Asc(endCol) - Asc("A") + 1 > Asc("G") - Asc("A") And (Asc(endCol) - Asc("A") + 1 <= Asc("P") - Asc("G") + Asc("G") - Asc("A") + 1))) Then
            'The calculations area is RESERVED for calc + clear cycles, so should not allow anything to be written _
            there in order to save the time and avoid increasing the complexity. Therefore no labels at the empty rows, _
            now it is more like whitespace in code!
            'If (ClearEmptyDataRowCSS) Then
            'for now only single row operations, sorry
            'If r.Row = endRow Then
            'If IsEmptyRange(Range("A" + r, Row + ":F" + endRow)) Then
            'IsFromTheInputRange = True
            'Exit Function
            'Else
            
            'End If
            'End If
            'End If
         Else
            IsFromTheInputRange = True
            Exit Function
         End If
      End If
   End If
   MsgBoxW ("Kolonas G-P tiek apr" + LV("e") + LV("k") + "in" + LV("a") + "tas, t" + LV("a") + "p" + LV("e") + "c " + LV("s") + "eit l" + LV("u") + "dzu neko nevad" + LV("i") + "t")
   IsFromTheInputRange = False
End Function

' Thank you: AgentRev ( https://stackoverflow.com/users/909968/agentrev )
' for this function: https://stackoverflow.com/a/31346246
Public Function IsEmptyRange(r As Range) As Boolean
   If IsObject(r) Then
      If Not r Is Nothing Then
         'delete operation yields an emty range!
         Dim myCell As Range
         For Each myCell In r
            If myCell.value <> "" Then
               IsEmptyRange = False
               Exit Function
            End If
         Next myCell
      End If
   End If
   IsEmptyRange = True
End Function

'IsEmptyRange is very slow when a column is being inserted therefore a quicker check
'0 neither or error, 1 row, 2 column, 4 multiple
Public Function IsColumOrRowBeingInsertedOrDeleted(ByVal r As Range) As Byte
   If IsObject(r) Then
      If Not r Is Nothing Then
         If (IndexOf(r.Address, ":") > -1) Then
            Dim ArrParts() As String
            ArrParts = Split(r.Address, ":")
            If (UBound(ArrParts) - LBound(ArrParts) + 1 = 2) Then
               Dim colStart As String, colEnd As String
               colStart = Mid(ArrParts(0), 2, Len(ArrParts(0)) - 1)
               colEnd = Mid(ArrParts(1), 2, Len(ArrParts(1)) - 1)
               'single row/column there never will be an address $a$1:$a$1, always reduced to $a$1
               If colStart = colEnd Then
                  If IsNumeric(colStart) Then
                     IsColumOrRowBeingInsertedOrDeleted = 1
                     Exit Function
                  Else
                     IsColumOrRowBeingInsertedOrDeleted = 2
                     Exit Function
                  End If
                  'multiple rows/columns
                ElseIf IndexOf(colStart, "$") = -1 Then
                  If IsNumeric(colStart) Then
                  'set both bits, row and multiple
                     IsColumOrRowBeingInsertedOrDeleted = 4 + 1
                     Exit Function
                  Else
                  'set both bits, column and multiple
                     IsColumOrRowBeingInsertedOrDeleted = 4 + 2
                     Exit Function
                  End If
               End If
            End If
         End If

      End If
   End If
   IsColumOrRowBeingInsertedOrDeleted = 0
End Function

Private Function HasCellValidation(cell As Range) As Boolean
   Dim t: t = Null

   On Error Resume Next
   t = cell.Validation.Type
   On Error GoTo 0

   HasCellValidation = Not IsNull(t)
End Function

' Thank you : Tushar Mehta
' for this function : http://www.tushar-mehta.com/publish_train/xl_vba_cases/1003%20MinMaxVals.shtml
Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft Excel") As VbMsgBoxResult
   MsgBoxW = MessageBoxW(Application.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

' Thank you : Citanaf ( https://stackoverflow.com/users/4383598/citanaf )
' for this function : https://stackoverflow.com/a/40411324/16769661
Sub ClearStatusBar()
Application.StatusBar = False
End Sub


Sub UnMergeAllMerged(r As Range)
   Dim myCell As Range
   Dim s As String
   s = ""
   If IsObject(r) Then
      If Not r Is Nothing Then
         For Each myCell In r
            If myCell.MergeCells Then
               myCell.MergeArea.UnMerge
               s = s + myCell.Address + ", "
            End If
         Next
         If Not s = "" Then
            s = "Neapvienot " + LV("s") + LV("u") + "nas! " + s
         Else
            s = LV("S") + LV("u") + "nas OK."
         End If
         If Application.StatusBar = False Then
            Application.StatusBar = Left(s, 255)
            Application.OnTime Now + TimeValue("00:00:02"), "ClearStatusBar"
         Else
            Application.StatusBar = CStr(Application.StatusBar) + ":::" + Left(s, 255 - 3 - Len(CStr(Application.StatusBar)))
         End If
      End If
   End If
End Sub

