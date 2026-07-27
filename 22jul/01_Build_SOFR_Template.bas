Attribute VB_Name = "SOFRTemplateBuilder"
Option Explicit

' ============================================================================
' SOFR INPUT TEMPLATE BUILDER - EXCEL VBA ONLY
'
' Import this module and 02_Run_SOFR_Simulation.bas into a blank .xlsm file.
' Run BuildSOFRTemplate once, paste the curve, then run RunSOFRSimulation.
'
' The builder preserves populated Curve, Config and FOMC sheets.
' ============================================================================


Public Sub BuildSOFRTemplate()

    Dim oldCalc As XlCalculation

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    SetupReadMeSheet
    SetupCurveSheet
    SetupConfigSheet
    SetupFOMCSheet

    ThisWorkbook.Worksheets("Read Me").Activate

    MsgBox "SOFR template created." & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Paste the curve into the Curve sheet." & vbCrLf & _
           "2. Review Config and FOMC." & vbCrLf & _
           "3. Run RunSOFRSimulation.", vbInformation

CleanExit:
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

CleanFail:
    MsgBox "Unable to build the SOFR template:" & vbCrLf & Err.Description, vbCritical
    Resume CleanExit

End Sub


' Compatibility alias.
Public Sub SetupSOFRWorkbook()
    BuildSOFRTemplate
End Sub


Private Sub SetupReadMeSheet()

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Read Me")
    ClearSheet ws

    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = "SOFR VBA Analysis Model"
    ApplyTitle ws.Range("A1:H1")

    ws.Range("A3").Value = "1. Run BuildSOFRTemplate once to create the input structure."
    ws.Range("A4").Value = "2. Paste the historical curve into the Curve sheet."
    ws.Range("A5").Value = "3. Update optional settings and starting dates in Config."
    ws.Range("A6").Value = "4. Update FOMC history when required."
    ws.Range("A7").Value = "5. Run RunSOFRSimulation to refresh all analysis sheets."

    ws.Range("A9").Value = "Required Curve headers"
    ws.Range("A10").Value = "Date | ON | 1M | 2M | 3M | 6M"

    ws.Range("A12").Value = "Rate formats accepted"
    ws.Range("A13").Value = "3.5500 or 0.0355. The simulation normalizes both to 3.55%."

    ws.Range("A15").Value = "Core methodology"
    ws.Range("A16").Value = "ACT/360; first available curve date on or after target maturity; interest reinvested at the same tenor on each rollover."

    ws.Range("A18").Value = "Local model"
    ws.Range("A19").Value = "No Python, internet connection, add-in or external file is required."

    AddRunButton ws

    ws.Range("A3:A19").WrapText = True
    ws.Columns("A").ColumnWidth = 110
    ws.Rows("3:19").RowHeight = 24
    ws.Range("A3:A19").Font.Name = "Aptos"
    ws.Range("A9,A12,A15,A18").Font.Bold = True
    ws.Range("A9,A12,A15,A18").Font.Color = RGB(23, 54, 93)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

End Sub


Private Sub AddRunButton(ByVal ws As Worksheet)

    Dim btn As Button

    On Error Resume Next
    ws.Buttons("btnRunSOFRSimulation").Delete
    On Error GoTo 0

    Set btn = ws.Buttons.Add(ws.Range("C3").Left, ws.Range("C3").Top, 190, 38)
    btn.Name = "btnRunSOFRSimulation"
    btn.Caption = "Run SOFR Simulation"
    btn.OnAction = "RunSOFRSimulation"
    btn.Font.Bold = True

End Sub


Private Sub SetupCurveSheet()

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Curve")

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        WriteRowValues ws.Range("A1"), Array("Date", "ON", "1M", "2M", "3M", "6M")
        ApplyHeader ws.Range("A1:F1")

        ws.Range("A2").Value = DateSerial(2024, 1, 2)
        WriteRowValues ws.Range("B2"), Array(5.33, 5.35, 5.36, 5.37, 5.4)

        ws.Range("A2:A10000").NumberFormat = "yyyy-mm-dd"
        ws.Range("B2:F10000").NumberFormat = "0.0000"
        ws.Range("A2:F10000").Font.Color = RGB(0, 0, 255)
        ws.Columns("A:F").ColumnWidth = 14
        ws.Rows(1).RowHeight = 28

        ws.Range("H1").Value = "Replace the sample row with your curve."
        ws.Range("H1").Font.Bold = True
        ws.Range("H2").Value = "Rates may be entered as 3.55 or 0.0355."
        ws.Range("H1:H2").WrapText = True
        ws.Columns("H").ColumnWidth = 42
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False

End Sub


Private Sub SetupConfigSheet()

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Config")

    If Len(Trim$(CStr(ws.Range("A1").Value))) > 0 Then Exit Sub

    ws.Range("A1:G1").Merge
    ws.Range("A1").Value = "SOFR Model Configuration"
    ApplyTitle ws.Range("A1:G1")

    WriteRowValues ws.Range("A3"), Array("Setting", "Value")
    ApplyHeader ws.Range("A3:B3")

    ws.Range("A4:A16").Value = Application.Transpose(Array( _
        "Analysis Start Date", _
        "Analysis End Date", _
        "Risk Start Date", _
        "Rising Regime Threshold (bp)", _
        "Falling Regime Threshold (bp)", _
        "Frontier Weight Step", _
        "Analysis Notional", _
        "Next FOMC Start Date", _
        "Next FOMC Decision Date", _
        "Next Policy Effective Date", _
        "Current Target Lower", _
        "Current Target Upper", _
        "Presentation Base Case"))

    ws.Range("B4").Value = ""
    ws.Range("B5").Value = ""
    ws.Range("B6").Value = DateSerial(2023, 1, 1)
    ws.Range("B7").Value = 5
    ws.Range("B8").Value = -5
    ws.Range("B9").Value = 0.05
    ws.Range("B10").Value = 100
    ws.Range("B11").Value = ""
    ws.Range("B12").Value = ""
    ws.Range("B13").Value = ""
    ws.Range("B14").Value = ""
    ws.Range("B15").Value = ""
    ws.Range("B16").Value = "Hold"

    ws.Range("A18").Value = "Curve Source"
    ws.Range("B18").Value = "User-provided local workbook"

    WriteRowValues ws.Range("D3"), Array("Starting Period Label", "Requested Start Date")
    ApplyHeader ws.Range("D3:E3")
    WriteJaggedArray ws.Range("D4"), Array( _
        Array("Jan-24", DateSerial(2024, 1, 1)), _
        Array("Jul-24", DateSerial(2024, 7, 1)), _
        Array("Jan-25", DateSerial(2025, 1, 1)), _
        Array("Jul-25", DateSerial(2025, 7, 1))), 2

    ws.Range("B4:B6,B11:B15,E4:E100").NumberFormat = "yyyy-mm-dd"
    ws.Range("B9").NumberFormat = "0%"
    ws.Range("B14:B15").NumberFormat = "0.00%"
    ws.Range("B4:B18,E4:E100").Font.Color = RGB(0, 0, 255)
    ws.Range("B4:B18,E4:E100").Interior.Color = RGB(255, 242, 204)

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 23
    ws.Columns("D").ColumnWidth = 23
    ws.Columns("E").ColumnWidth = 21
    ws.Columns("G").ColumnWidth = 42

    ws.Range("G3").Value = "Leave analysis start/end dates blank to use the available curve period."
    ws.Range("G3").WrapText = True

    ws.Activate
    ActiveWindow.DisplayGridlines = False

End Sub


Private Sub SetupFOMCSheet()

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("FOMC")

    If Len(Trim$(CStr(ws.Range("A1").Value))) > 0 Then Exit Sub

    WriteRowValues ws.Range("A1"), Array("Meeting Date", "Decision", "Action (bp)", "SEP")
    ApplyHeader ws.Range("A1:D1")

    AddDefaultFOMCRow ws, 2, DateSerial(2023, 2, 1), "Hike", 25, "No"
    AddDefaultFOMCRow ws, 3, DateSerial(2023, 3, 22), "Hike", 25, "Yes"
    AddDefaultFOMCRow ws, 4, DateSerial(2023, 5, 3), "Hike", 25, "No"
    AddDefaultFOMCRow ws, 5, DateSerial(2023, 6, 14), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 6, DateSerial(2023, 7, 26), "Hike", 25, "No"
    AddDefaultFOMCRow ws, 7, DateSerial(2023, 9, 20), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 8, DateSerial(2023, 11, 1), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 9, DateSerial(2023, 12, 13), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 10, DateSerial(2024, 1, 31), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 11, DateSerial(2024, 3, 20), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 12, DateSerial(2024, 5, 1), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 13, DateSerial(2024, 6, 12), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 14, DateSerial(2024, 7, 31), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 15, DateSerial(2024, 9, 18), "Cut", -50, "Yes"
    AddDefaultFOMCRow ws, 16, DateSerial(2024, 11, 7), "Cut", -25, "No"
    AddDefaultFOMCRow ws, 17, DateSerial(2024, 12, 18), "Cut", -25, "Yes"
    AddDefaultFOMCRow ws, 18, DateSerial(2025, 1, 29), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 19, DateSerial(2025, 3, 19), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 20, DateSerial(2025, 5, 7), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 21, DateSerial(2025, 6, 18), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 22, DateSerial(2025, 7, 30), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 23, DateSerial(2025, 9, 17), "Cut", -25, "Yes"
    AddDefaultFOMCRow ws, 24, DateSerial(2025, 10, 29), "Cut", -25, "No"
    AddDefaultFOMCRow ws, 25, DateSerial(2025, 12, 10), "Cut", -25, "Yes"
    AddDefaultFOMCRow ws, 26, DateSerial(2026, 1, 28), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 27, DateSerial(2026, 3, 18), "Hold", 0, "Yes"
    AddDefaultFOMCRow ws, 28, DateSerial(2026, 4, 29), "Hold", 0, "No"
    AddDefaultFOMCRow ws, 29, DateSerial(2026, 6, 17), "Hold", 0, "Yes"

    ws.Range("A2:A1000").NumberFormat = "yyyy-mm-dd"
    ws.Range("A2:D1000").Font.Color = RGB(0, 0, 255)
    ws.Columns("A:D").ColumnWidth = 17

    ws.Range("F1").Value = "Add future meetings with a blank Decision if required."
    ws.Range("F1").Font.Bold = True
    ws.Columns("F").ColumnWidth = 48

    ws.Activate
    ActiveWindow.DisplayGridlines = False

End Sub


Private Sub AddDefaultFOMCRow(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                              ByVal meetingDate As Date, ByVal decision As String, _
                              ByVal actionBp As Double, ByVal sepValue As String)

    ws.Cells(rowNumber, 1).Value = meetingDate
    ws.Cells(rowNumber, 2).Value = decision
    ws.Cells(rowNumber, 3).Value = actionBp
    ws.Cells(rowNumber, 4).Value = sepValue

End Sub


Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function


Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If

End Function


Private Sub ClearSheet(ByVal ws As Worksheet)

    Dim i As Long

    On Error Resume Next
    ws.Cells.UnMerge
    ws.Cells.Clear

    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    On Error GoTo 0

End Sub


Private Sub ApplyTitle(ByVal target As Range)

    With target
        .Interior.Color = RGB(23, 54, 93)
        .Font.Name = "Aptos"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With

End Sub


Private Sub ApplyHeader(ByVal target As Range)

    With target
        .Interior.Color = RGB(217, 234, 247)
        .Font.Name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 36
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(183, 201, 214)
        .Borders.Weight = xlThin
    End With

End Sub


Private Sub WriteRowValues(ByVal startCell As Range, ByVal valuesArray As Variant)

    Dim output() As Variant
    Dim i As Long
    Dim countValue As Long

    countValue = UBound(valuesArray) - LBound(valuesArray) + 1
    ReDim output(1 To 1, 1 To countValue)

    For i = LBound(valuesArray) To UBound(valuesArray)
        output(1, i - LBound(valuesArray) + 1) = valuesArray(i)
    Next i

    startCell.Resize(1, countValue).Value = output

End Sub


Private Sub WriteJaggedArray(ByVal startCell As Range, ByVal data As Variant, _
                             ByVal columnCount As Long)

    Dim output() As Variant
    Dim r As Long
    Dim c As Long

    ReDim output(1 To UBound(data) - LBound(data) + 1, 1 To columnCount)

    For r = LBound(data) To UBound(data)
        For c = 0 To columnCount - 1
            output(r - LBound(data) + 1, c + 1) = data(r)(c)
        Next c
    Next r

    startCell.Resize(UBound(output, 1), UBound(output, 2)).Value = output

End Sub
