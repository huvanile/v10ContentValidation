Option Explicit

'NOTE: This sub contains some random scripts that were written to support the performance of a few very targeted validation tests for v10


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
    Const CHECKMIN As Integer = 6
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

