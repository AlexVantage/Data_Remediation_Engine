' DefectIntroducer.vba
' Purpose: Generates a teaching-grade dirty dataset for the Diagnostics & Rebuild template.
' This script introduces defects — it is NOT the cleaning solution.
' See 01_Diagnostics_Rebuild_Template_v1.xlsx for the repair formulas.
Sub IntroduceDefects()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim rng As Range
    Dim i As Long
    Dim r As Double
    
    ' --- 1. DUPLICATE ROWS (15 duplicates) ---
    Dim dupeCount As Integer
    dupeCount = 0
    Dim targetRow As Long
    targetRow = lastRow + 1
    
    Do While dupeCount < 15
        Dim srcRow As Long
        srcRow = Int(Rnd * (lastRow - 1)) + 2
        ws.Rows(srcRow).Copy ws.Rows(targetRow)
        targetRow = targetRow + 1
        dupeCount = dupeCount + 1
    Loop
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' --- 2. CUSTOMER_NAME: leading/trailing whitespace + lowercase (30% of rows) ---
    For i = 2 To lastRow
        r = Rnd
        If r < 0.15 Then
            ws.Cells(i, 2).Value = " " & ws.Cells(i, 2).Value   ' leading space
        ElseIf r < 0.30 Then
            ws.Cells(i, 2).Value = ws.Cells(i, 2).Value & " "   ' trailing space
        End If
        If Rnd < 0.15 Then
            ws.Cells(i, 2).Value = LCase(ws.Cells(i, 2).Value)  ' random lowercase
        End If
    Next i
    
    ' --- 3. CUSTOMER_EMAIL: mixed case ---
    For i = 2 To lastRow
        If Rnd < 0.5 Then
            ws.Cells(i, 3).Value = UCase(ws.Cells(i, 3).Value)
        End If
    Next i
    
    ' --- 4. PHONE_NUMBER: inconsistent formats ---
    Dim phoneFormats(3) As String
    phoneFormats(0) = "(###) ###-####"
    phoneFormats(1) = "###-###-####"
    phoneFormats(2) = "##########"
    phoneFormats(3) = "###.###.####"
    
    For i = 2 To lastRow
        Dim rawPhone As String
        rawPhone = ws.Cells(i, 4).Value
        ' Strip to digits only first
        Dim digits As String
        digits = ""
        Dim c As String
        Dim j As Integer
        For j = 1 To Len(rawPhone)
            c = Mid(rawPhone, j, 1)
            If c >= "0" And c <= "9" Then digits = digits & c
        Next j
        If Len(digits) >= 10 Then
            digits = Right(digits, 10)
            Dim fmt As Integer
            fmt = Int(Rnd * 4)
            Select Case fmt
                Case 0: ws.Cells(i, 4).Value = "(" & Left(digits, 3) & ") " & Mid(digits, 4, 3) & "-" & Right(digits, 4)
                Case 1: ws.Cells(i, 4).Value = Left(digits, 3) & "-" & Mid(digits, 4, 3) & "-" & Right(digits, 4)
                Case 2: ws.Cells(i, 4).Value = digits
                Case 3: ws.Cells(i, 4).Value = Left(digits, 3) & "." & Mid(digits, 4, 3) & "." & Right(digits, 4)
            End Select
        End If
    Next i
    
    ' --- 5. ORDER_DATE: convert 5% to text-formatted dates ---
    For i = 2 To lastRow
        If Rnd < 0.05 Then
            ws.Cells(i, 5).Value = "'" & Format(ws.Cells(i, 5).Value, "MM/DD/YYYY")
        End If
    Next i
    
    ' --- 6. REGION: 5% blanks ---
    For i = 2 To lastRow
        If Rnd < 0.05 Then ws.Cells(i, 7).Value = ""
    Next i
    
    ' --- 7. REVENUE_USD: 5% blanks ---
    For i = 2 To lastRow
        If Rnd < 0.05 Then ws.Cells(i, 10).Value = ""
    Next i
    
    ' --- 8. QUANTITY: 5% blanks ---
    For i = 2 To lastRow
        If Rnd < 0.05 Then ws.Cells(i, 11).Value = ""
    Next i
    
    MsgBox "Defects introduced. Row count: " & lastRow & ". Save as shopify_dump_dirty.xlsx."
End Sub