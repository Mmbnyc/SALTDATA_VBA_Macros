Attribute VB_Name = "A_ProDSS_Kor404_Final"
Option Explicit

'==================================================
' MAIN CONTROLLER
'==================================================
Sub Populate_ProDSS()

    Dim NumYear As Integer, NameMonth As String, intMonth As Integer
    Dim BasePath As String, MonthPath As String
    Dim SaveFileName As String
    Dim TemplateWB As Workbook
    Dim ProDSSwb As Workbook, ProDSSws As Worksheet
    Dim MeanRows As Collection
    Dim i As Long, startRow As Long, endRow As Long
    Dim sheetName As String
    Dim newWb As Workbook, newWs As Worksheet
    Dim lastRow As Long
    Dim isQC As Boolean, isCCV As Boolean, isMW As Boolean
    Dim f As String
    Dim srcWb As Workbook
    Dim KorFile As String
    Dim qcCounter As Long, ccvCounter As Long
    
    Dim sampleDate As String
    Dim dateArray() As String
    Dim arraySize As Integer
    
    Dim timeAscending As Boolean
    

    '--------------------------------------------------
    ' Visual feedback (ON by request)
    '--------------------------------------------------
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Starting ProDSS processing..."

    If Not ConfirmFolderStructure Then GoTo Cleanup
    GetUserInput NumYear, NameMonth, intMonth

    BasePath = Environ("OneDriveCommercial") & "\Monitoring Wells\Chloride monitoring\"
    MonthPath = BasePath & NumYear & "\" & NameMonth & "\"

    If Dir(MonthPath, vbDirectory) = "" Then
        MsgBox "Missing folder: " & MonthPath, vbCritical
        GoTo Cleanup
    End If

    SaveFileName = MonthPath & Format(DateSerial(NumYear, intMonth, 1), "mmyyyy") & ".xlsm"
    ThisWorkbook.SaveCopyAs SaveFileName
    Set TemplateWB = Workbooks.Open(SaveFileName)

    KorFile = Dir(MonthPath & "Kor Measurement File Export*.xlsx")
    If KorFile = "" Then
        MsgBox "Kor Measurement export not found.", vbCritical
        GoTo Cleanup
    End If

    Set ProDSSwb = Workbooks.Open(MonthPath & KorFile)
    Set ProDSSws = GetKorExportSheet(ProDSSwb)
    
    timeAscending = IsTimeAscending(ProDSSws)

    lastRow = ProDSSws.Cells(ProDSSws.rows.Count, "A").End(xlUp).Row
    Set MeanRows = GetMeanRows(ProDSSws, lastRow)

    qcCounter = 1
    ccvCounter = 1
    
    arraySize = 0
    
    '==================================================
    ' PROCESS EACH MEAN VALUE BLOCK
    '==================================================
    For i = 1 To MeanRows.Count

        Application.StatusBar = "Processing block " & i & " of " & MeanRows.Count

        startRow = MeanRows(i)
        
    ''''
        'endRow = IIf(i < MeanRows.Count, MeanRows(i + 1) - 2, lastRow)
        If i < MeanRows.Count Then
            endRow = MeanRows(i + 1) - 2
        Else
            endRow = lastRow
        End If
    ''''

        sheetName = Trim(ProDSSws.Cells(startRow + 6, "C").Value)
        If sheetName = "" Then sheetName = "UNNAMED_" & i

        sheetName = UCase(Left(Replace(sheetName, ":", "_"), 31))

        isQC = (sheetName Like "QC*")
        isCCV = (sheetName Like "CCV*")
        isMW = Not (isQC Or isCCV)

        If isQC Then
            sheetName = "QC" & qcCounter
            qcCounter = qcCounter + 1
        ElseIf isCCV Then
            sheetName = "CCV" & ccvCounter
            ccvCounter = ccvCounter + 1
        End If

        Set newWb = Workbooks.Add(xlWBATWorksheet)
        Set newWs = newWb.Sheets(1)
        newWs.name = sheetName

        ProDSSws.Range("A" & startRow & ":K" & endRow).Copy newWs.Range("A1")
        newWs.rows("1:4").Delete
        
'''''''''''''''''''''''''''''''testing sample date code
        'Dim n As Integer
        'n = 0
        'If isMW Then
        sampleDate = newWs.Cells(2, "A").Value
        Debug.Print sampleDate
        newWs.Range("Z1").Value = sampleDate
                
            'ReDim Preserve dateArray(0 To arraySize)
            'dateArray(arraySize) = sampleDate
            'arraySize = arraySize + 1
        'End If
        
        '----------------------------------------------
        ' Reverse #1: remove upward readings (domain)
        '----------------------------------------------
        
        If Not timeAscending Then
            ReverseImportedData newWs
        End If

        CleanColumns newWs.Cells(1, 1).CurrentRegion, (isQC Or isCCV)
        ArrangeColumns newWs

        If isMW Then FilterEvery10FeetSheet newWs

        '----------------------------------------------
        ' Reverse #2: format for table (presentation)
        '----------------------------------------------
        If isMW Then ReverseForTable newWs

        newWb.SaveAs MonthPath & sheetName & ".xlsx"
        newWb.Close False
    Next i

    ProDSSwb.Close False

'    Dim element As Variant
'    For Each element In dateArray
'        Debug.Print element
'    Next element

    '==================================================
    ' COPY ALL SHEETS INTO TEMPLATE
    '==================================================
    Application.StatusBar = "Assembling template workbook..."

    f = Dir(MonthPath & "*.xlsx")
    Do While f <> ""
        If InStr(1, f, "Kor", vbTextCompare) = 0 Then
            Set srcWb = Workbooks.Open(MonthPath & f)

            srcWb.Sheets(1).Copy After:=TemplateWB.Sheets("Table")

            If Left$(UCase(srcWb.Sheets(1).name), 2) = "MW" Then
                CopyMWDataToTable srcWb.Sheets(1), TemplateWB.Sheets("Table"), sampleDate
            End If

            srcWb.Close False
        End If
        f = Dir
    Loop

    ArrangeMWAndQCTabs TemplateWB

    MsgBox "ProDSS processing complete.", vbInformation

Cleanup:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

'==================================================
' SUPPORT FUNCTIONS
'==================================================
Sub ReverseImportedData(ws As Worksheet)
    ReverseRange ws, 2
End Sub

Sub ReverseForTable(ws As Worksheet)
    ReverseRange ws, 2
End Sub

Private Sub ReverseRange(ws As Worksheet, firstRow As Long)
    Dim lastRow As Long, i As Long
    Dim t1, t2
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For i = 0 To (lastRow - firstRow) \ 2
        t1 = ws.Cells(firstRow + i, 1).Value
        t2 = ws.Cells(firstRow + i, 2).Value
        ws.Cells(firstRow + i, 1).Value = ws.Cells(lastRow - i, 1).Value
        ws.Cells(firstRow + i, 2).Value = ws.Cells(lastRow - i, 2).Value
        ws.Cells(lastRow - i, 1).Value = t1
        ws.Cells(lastRow - i, 2).Value = t2
    Next i
End Sub

Sub FilterEvery10FeetSheet(ws As Worksheet)
    Dim i As Long, lastRow As Long, nextDepth As Double
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    nextDepth = 10
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value < nextDepth Then
            ws.rows(i).Delete
        Else
            nextDepth = nextDepth + 10
        End If
    Next i
End Sub

'Sub CopyMWDataToTable(wsMW As Worksheet, wsTable As Worksheet)
'    Dim mwNum As Long, targetCol As Long, lastRow As Long
'    mwNum = Val(ExtractMWNumberFromName(wsMW.name))
'    If mwNum = 0 Then Exit Sub
'    targetCol = mwNum + 1
'    lastRow = wsMW.Cells(wsMW.rows.Count, 2).End(xlUp).Row
'    wsTable.Cells(6, targetCol).Resize(lastRow - 1).Value = _
'        wsMW.Range("B2:B" & lastRow).Value
'End Sub

Sub CopyMWDataToTable(wsMW As Worksheet, wsTable As Worksheet, ByVal sampleDate As String)

    Dim mwNum As Long
    Dim targetCol As Long
    Dim lastRow As Long
    
    Dim sampleDateZQ As String
    
    'Dim sampleDate As Variant

    '----------------------------------------------
    ' Determine MW number
    '----------------------------------------------
    mwNum = Val(ExtractMWNumberFromName(wsMW.name))
    If mwNum = 0 Then Exit Sub

    targetCol = mwNum + 1   ' Column B = MW1

    '----------------------------------------------
    ' Capture sample date (Column A, consistent)
    '----------------------------------------------
    
    sampleDateZQ = wsMW.Range("Q1").Value

    ' Write date to Table row 4
    wsTable.Cells(4, targetCol).Value = sampleDateZQ

    '----------------------------------------------
    ' Copy filtered MW data
    '----------------------------------------------
    lastRow = wsMW.Cells(wsMW.rows.Count, 2).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    wsTable.Cells(6, targetCol).Resize(lastRow - 1).Value = _
        wsMW.Range("B2:B" & lastRow).Value

End Sub


Function ExtractMWNumberFromName(wsName As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(wsName)
        If Mid(wsName, i, 1) Like "#" Then s = s & Mid(wsName, i, 1)
    Next i
    ExtractMWNumberFromName = s
End Function

'==================================================
' (Other helpers unchanged)
'==================================================
Function ConfirmFolderStructure() As Boolean
    ConfirmFolderStructure = (MsgBox( _
        "Ensure folder structure exists:" & vbCrLf & _
        "Monitoring Wells\Chloride monitoring\YEAR\MONTH", _
        vbOKCancel, "Confirm") = vbOK)
End Function
Sub GetUserInput(ByRef NumYear As Integer, ByRef NameMonth As String, ByRef intMonth As Integer)
    NumYear = Val(InputBox("Enter year:", , Year(Date)))
    NameMonth = LCase(InputBox("Enter month (mmm):", , Format(Date, "mmm")))
    intMonth = Month(DateValue("01-" & NameMonth & "-" & NumYear))
End Sub
Function GetKorExportSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase(ws.name) Like "kor measurement file export*" Then
            Set GetKorExportSheet = ws
            Exit Function
        End If
    Next ws
End Function
Function GetMeanRows(ws As Worksheet, lastRow As Long) As Collection
    Dim c As New Collection, i As Long
    For i = 1 To lastRow
        If Trim(ws.Cells(i, "A").Value) = "MEAN VALUE:" Then c.Add i
    Next i
    Set GetMeanRows = c
End Function
Sub CleanColumns(rng As Range, skipClean As Boolean)
    Dim c As Long
    If skipClean Then Exit Sub
    For c = rng.Columns.Count To 1 Step -1
        Select Case UCase(Trim(rng.Cells(1, c).Value))
            Case "DEPTH FT", "SPCOND µS/CM"
            Case "SPCOND µS/CM", "DEPTH FT"
            Case Else
                rng.Columns(c).Delete
        End Select
    Next c
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Sub ArrangeColumns(ws As Worksheet)
'    Dim c As Long
'    For c = ws.UsedRange.Columns.Count To 1 Step -1
'        If UCase(Trim(ws.Cells(1, c).Value)) = "DEPTH FT" Then
'            ws.Columns(c).Cut
'            ws.Columns(1).Insert
'            Application.CutCopyMode = False
'            Exit For
'        End If
'    Next c
'End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub ArrangeColumns(ws As Worksheet)
    Dim c As Long
    Dim depthCol As Long
    
    depthCol = 0
    
    '--------------------------------------------------
    ' Find DEPTH FT column
    '--------------------------------------------------
    For c = 1 To ws.UsedRange.Columns.Count
        If UCase(Trim(ws.Cells(1, c).Value)) = "DEPTH FT" _
            Or UCase(Trim(ws.Cells(1, c).Value)) = "Pressure (Ft H2O)" Then
            depthCol = c
            Exit For
        End If
    Next c
    '--------------------------------------------------
    ' Exit safely if not found or already in position
    '--------------------------------------------------
    If depthCol = 0 Or depthCol = 1 Then Exit Sub

    '--------------------------------------------------
    ' Move DEPTH FT to column A
    '--------------------------------------------------
    ws.Columns(depthCol).Cut
    ws.Columns(1).Insert
    Application.CutCopyMode = False
End Sub
    
'    Dim col As Integer
'    For col = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
'        If UCase(Trim(ActiveSheet.Cells(1, col).Value)) = "Pressure (Ft H2O)" _
'           Or UCase(Trim(ActiveSheet.Cells(1, col).Value)) = "DEPTH FT" Then
'            ActiveSheet.Columns(col).Cut
'            ActiveSheet.Columns(1).Insert Shift:=xlToRight
'            Application.CutCopyMode = False ' Clear clipboard
'            Exit For
'        End If
'    Next col
'End Sub

Sub ArrangeMWAndQCTabs(wb As Workbook)
    Dim i As Long, ws As Worksheet
    Dim tableIndex As Long
    
    tableIndex = wb.Sheets("Table").Index
    
    ' MW1ñMW10
    For i = 1 To 10
        On Error Resume Next
        Set ws = wb.Sheets("MW" & i)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Move After:=wb.Sheets(tableIndex)
            tableIndex = ws.Index
        End If
        Set ws = Nothing
    Next i
    
    ' QC1ñQC10
    For i = 1 To 10
        On Error Resume Next
        Set ws = wb.Sheets("QC" & i)
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Move After:=wb.Sheets(tableIndex)
            tableIndex = ws.Index
        End If
        Set ws = Nothing
    Next i
End Sub

Function IsTimeAscending(ws As Worksheet) As Boolean
    Dim t1 As Date, t2 As Date
    
    On Error GoTo FailSafe
    
    t1 = ws.Cells(11, "B").Value
    t2 = ws.Cells(12, "B").Value
    
    IsTimeAscending = (t2 >= t1)
    Exit Function

FailSafe:
    ' Default to ascending if anything unexpected happens
    IsTimeAscending = True
End Function






  Ôæ$XLa≈XsX.   7                  Ãˆ S S C O N   I n s p e c t i o n s    j 1     øXÀv0 2024IN~1  R 	  Ôæ$X¸l≈XsX.   É   =               C2 0 2 4   I n s p e c t i o n s    b 1     ƒXpá FATVIL~2  J 	  ÔæsX™å≈XsX.   /«                  %}f a t   v i l   e a s t        §ãâ  Ä P‡O– Í:i¢ÿ +00ù /C:\                   x 1     êXŸ> Users d 	  ÔæßT,* Xçm.   *¥   	         :     Ñ´U s e r s   @ s h e l l 3 2 . d l l , - 2 1 8 1 3    \ 1     ∆Xhg MICHAE~1  D 	  ÔæiWz XSm.   x1   