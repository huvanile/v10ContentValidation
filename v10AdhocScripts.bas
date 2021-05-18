Option Explicit

'NOTE: This sub contains some random scripts that were written to support the performance of a few very targeted validation tests for v10

'PURPOSE: deep cleans all requirement statements in in the selection
Sub deepCleanSelection()
    Dim eCellCount  As Integer: eCellCount = 0
    Dim cell As Range
    For Each cell In Selection
        Debug.Print cell.Row
        cell.Value = cleanRS(cell.Value, True, True, True, True)
        If cell.Value = "" Then eCellCount = eCellCount + 1
        If eCellCount > 10 Then End
    Next cell
End Sub

'PURPOSE: strip element counts from all requirements in the selection
Sub stripElementCountsFromSelection()
    Dim eCellCount  As Integer: eCellCount = 0
    Dim cell As Range
    For Each cell In Selection
        cell.Value = stripElementCounts(cell.Value)
        cell.Value = stripElementCounts(cell.Value)
        If cell.Value = "" Then eCellCount = eCellCount + 1
        If eCellCount > 10 Then End
    Next cell
End Sub

'PURPOSE: Removes element counts from requirement statements.
    'ACCEPTS: A requirement statement
    'RETURNS: A requirement statement without any element counts present
Public Function stripElementCounts(str As String)
    Dim i As Integer
    Dim tmp As String: tmp = str
    For i = 1 To 20
        tmp = Replace(tmp, i & ")", "")
        tmp = Replace(tmp, "  ", " ")
    Next i
    stripElementCounts = tmp
End Function

'PURPOSE: Flags any element lists in requirements that are OR lists, using approach "B"
'NOTE: Any findings are added to column K of the active worksheet. It will NOT erase column K before it runs the checks.
'EXPECTED COLUMNS:
    '- column A: RID
    '- column G: requirement element count
    '- column H: requirement statement
Sub flagORListsApproachB()
    Dim cell As Range
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cleanRS(cell.Value)
        For i = 1 To target + 1
            Application.StatusBar = "Checking for short elements within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" Then
                If LCase(RS) Like "*" & i & ") or*" _
                Or LCase(RS) Like "*" & i & ". or*" Then
                    If Not LCase(RS) Like "*" & i & ") org*" Then addToK cell.Row, "or list: " & i
                End If
            End If
        Next i
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Flags any element lists in requirements that are OR lists, using approach "A"
'NOTE: Any findings are added to column K of the active worksheet. It will NOT erase column K before it runs the checks.
'EXPECTED COLUMNS:
    '- column A: RID
    '- column G: requirement element count
    '- column H: requirement statement
Sub flagORListsApproachA()
    Dim cell As Range
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cleanRS(cell.Value)
        For i = 1 To target + 1
            Application.StatusBar = "Checking for short elements within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" Then
                If LCase(RS) Like "* or " & i & ")*" _
                Or LCase(RS) Like "* or (" & i & ")*" _
                Or LCase(RS) Like "*\or " & i & ")*" _
                Or LCase(RS) Like "*\or (" & i & ")*" _
                Or LCase(RS) Like "*/or " & i & ")*" _
                Or LCase(RS) Like "*/or (" & i & ")*" Then
                    If Not LCase(RS) Like "*for " & i & "*" Then addToK cell.Row, "or list: " & i
                End If
            End If
        Next i
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Flags any elements present within requirement statements that are between CHECKLEN and CHECKMIN characters, written to find bad element numbering essentially
'NOTE: Any findings are added to column K of the active worksheet. It will NOT erase column K before it runs the checks.
'EXPECTED COLUMNS:
    '- column A: RID
    '- column G: requirement element count
    '- column H: requirement statement
Sub checkForShortElements()
    Dim cell As Range
    Const CHECKLEN As Integer = 2
    Const CHECKMIN As Integer = 10
    Const errText As String = "Element is between " & CHECKMIN & " and " & CHECKLEN & " characters: "
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cleanRS(cell.Value)
        For i = 1 To target + 1
            Application.StatusBar = "Checking for short elements within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" Then
                Dim splitholder: splitholder = Split(RS, ")")
                Dim s As Integer
                For s = LBound(splitholder) To UBound(splitholder)
                    If Len(Trim(splitholder(s))) <= CHECKLEN And Len(Trim(splitholder(s))) > CHECKMIN Then
                        addToK cell.Row, errText & splitholder(s)
                    End If
                Next s
            End If
        Next i
    Next cell
    Application.StatusBar = False
End Sub

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
    Debug.Print r
End Sub
