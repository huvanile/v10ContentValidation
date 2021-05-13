Option Explicit

'PURPOSE: This is the factory. Run this to run all of the RID checks at once.
'NOTE: Any findings are added to column K of the active worksheet. It will erase column K before it runs the checks.
'EXPECTED COLUMNS:
    '- column A: RID
    '- column G: requirement element count
    '- column H: requirement statement
Sub doRIDChecks()
    Range("k2:k" & lastUsedRow).ClearContents
    reconElementCounts
    checkRIDCategoryCode
    checkRIDDotsAndHyphens
    MsgBox "Done!"
End Sub

'PURPOSE: Compares the element counts as identified in the RID to the element numbers present in the element count column
Sub reconElementCounts()
    Dim cell As Range
    For Each cell In Range("a2:a" & lastUsedRow)
        Application.StatusBar = "Comparing RID's element count to stated element count | On row " & cell.Row
        Dim splitholder: splitholder = Split(cell.Value, ".")
        Dim ridElementCount As Integer: ridElementCount = CInt(Trim(splitholder(UBound(splitholder))))
        Dim statedElementCount As Integer: statedElementCount = CInt(Trim(Range("h" & cell.Row)))
        If ridElementCount <> statedElementCount Then
            addToK cell.Row, "RID's element count (numbers after last dot) don't reconcile to stated element count"
        End If
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Inspects the RIDs and alerts when less than 4 dots or less than 2 hyphens are present
'NOTE: Skips the contra requirements
Sub checkRIDDotsAndHyphens()
    Dim cell As Range
    For Each cell In Range("a2:a" & lastUsedRow)
        Application.StatusBar = "Checking RID hypen and dot count | On row " & cell.Row
        If Not cell.Value Like "*CONTRA*" Then
            If Not Len(cell.Value) - Len(Replace(cell.Value, ".", "")) = 4 Then
                addToK cell.Row, "RID dot count is off"
            End If
            If Not Len(cell.Value) - Len(Replace(cell.Value, "-", "")) = 2 Then
                addToK cell.Row, "RID hyphen count is off"
            End If
        End If
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Inspects the RIDs and alerts if the first 3 letters are not a valid category code
'NOTE: Skips the contra requirements
Sub checkRIDCategoryCode()
    Dim cell As Range
    For Each cell In Range("a2:a" & lastUsedRow)
        Application.StatusBar = "Checking RID category codes | On row " & cell.Row
        If Not cell.Value Like "*CONTRA*" Then
            If Not cell.Value Like "ASM*" _
            And Not cell.Value Like "BCM*" _
            And Not cell.Value Like "CCP*" _
            And Not cell.Value Like "CMP*" _
            And Not cell.Value Like "EIM*" _
            And Not cell.Value Like "IAM*" _
            And Not cell.Value Like "ITD*" _
            And Not cell.Value Like "OPS*" _
            And Not cell.Value Like "PEP*" _
            And Not cell.Value Like "PGM*" _
            And Not cell.Value Like "PIR*" _
            And Not cell.Value Like "SCS*" _
            And Not cell.Value Like "SLC*" _
            And Not cell.Value Like "TPR*" _
            And Not cell.Value Like "WFS*" Then
                addToK cell.Row, "RID doesn't start with a valid three-letter category code"
            End If
        End If
    Next cell
    Application.StatusBar = False
End Sub

'~~~~~~~ SUPPORTING SUBROUTINES AND FUNCTIONS ~~~~~~~~
'PURPOSE: This simply adds any findings to column K
Private Sub addToK(ByVal r As Integer, msg As String)
    Dim wb As String: wb = ActiveWorkbook.name
    Dim ws As String: ws = ActiveSheet.name
    If Workbooks(wb).Sheets(ws).Range("k1") = "" Then
        With Workbooks(wb).Sheets(ws).Range("k1")
            .FormulaLocal = "=CONCATENATE(COUNTif(K2:K9999,""*•*""),"" Possible Quality Issue(s)"")"
            .Font.Color = 255
            .Font.Bold = True
            .Font.Underline = True
            .EntireColumn.ColumnWidth = 70
        End With
    End If
    With Workbooks(wb).Sheets(ws).Range("k" & r)
        .WrapText = True
        If .Value = "" Then
            .Value = " • " & msg
        Else
            If InStr(1, Workbooks(wb).Sheets(ws).Range("k" & r).Value, msg) = 0 Then .Value = .Value & vbCrLf & " • " & msg
        End If
        
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlLeft
        .Font.Color = 255
                
    End With
    Workbooks(wb).Sheets(ws).Tab.Color = 255
    ReDim foundWord(0) As String
End Sub
