Option Explicit

'PURPOSE: This is the factory. Run this to run all of the RS checks at once.
'NOTE: Any findings are added to column K of the active worksheet. It will erase column K before it runs the checks.
'EXPECTED COLUMNS:
    '- column A: RID
    '- column G: requirement element count
    '- column H: requirement statement
Sub doRSChecks()
    Range("k2:k" & lastUsedRow).ClearContents
    checkForMissingNumbers
    checkForExtraNumbers
    checkForDupeNumbers
    MsgBox "Done!"
End Sub

'PURPOSE: Checks for duplicated element numbers present in a requirement statement
'NOTE: Skips the contra requirements
Sub checkForDupeNumbers()
    Dim cell As Range
    Const errText As String = "Requirement contains a duplicated number: "
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cell.Value
        For i = 1 To target + 1
            Application.StatusBar = "Checking for duplicate numbers within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" Then
                If RS Like "*" & i & ")*" _
                And Not RS Like "*." & i & ")*" _
                And Not RS Like "*(" & i & ")*" Then
                     If i < 10 Then
                        If Len(RS) - Len(Replace(RS, i & ")", "")) <> 2 Then
                            addToK cell.Row, errText & i
                        End If
                    ElseIf Len(RS) - Len(Replace(RS, i & ")", "")) <> 3 Then
                        addToK cell.Row, errText & i
                    End If
                ElseIf RS Like "*" & i & "-*" Then
                     If i < 10 Then
                        If Len(RS) - Len(Replace(RS, i & "-", "")) <> 2 Then
                            addToK cell.Row, errText & i
                        End If
                    ElseIf Len(RS) - Len(Replace(RS, i & ")", "")) <> 3 Then
                        addToK cell.Row, errText & i
                    End If
                End If
            End If
        Next i
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Checks for extra element numbers present in a requirement statement
'NOTE: Skips the contra requirements
Sub checkForExtraNumbers()
    Dim cell As Range
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cell.Value
        For i = target + 1 To 20
            Application.StatusBar = "Checking for EXTRA numbers within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" Then
                 If RS Like "*" & i & ")*" _
                 Or RS Like "*" & i & "-*" Then
                    If Not RS Like "*." & i & ")*" _
                    And Not RS Like "*(" & i & ")*" Then
                        addToK cell.Row, "Unexpected number found within requirement statement (as compared to the requirement's stated element count): " & i
                    End If
                End If
            End If
        Next i
    Next cell
    Application.StatusBar = False
End Sub

'PURPOSE: Checks for missing element numbers present in a requirement statement
'NOTE: Skips the contra requirements
Sub checkForMissingNumbers()
    Dim cell As Range
    For Each cell In Range("h2:h" & lastUsedRow)
        Dim target As Integer: target = CInt(Range("g" & cell.Row).Value)
        Dim i As Integer
        Dim RS As String: RS = cell.Value
        For i = 1 To target
            Application.StatusBar = "Checking for MISSING numbers  within RS | On row " & cell.Row & "." & i
            If Not Range("a" & cell.Row) Like "*CONTRA*" _
            And Not RS Like "*" & i & ")*" _
            And Not RS Like "*(356)*" _
            And Not RS Like "*" & i & ") days*" _
            And Not RS Like "*" & i & ".*" Then
                addToK cell.Row, "Cannot find expected number within requirement statement: " & i
            End If
        Next i
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
    Debug.Print r
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
End Function

'PURPOSE: Removes parenthetical clauses from requirements. It also runs the requirement through the cleanRS function described below
    'ACCEPTS: A requirement statement
    'RETURNS: A requirement statement without any parenthetical clauses present
Public Function cleanRSParens(theRS As String)
    cleanRSParens = cleanRS(theRS, True, True, False, True)
End Function

'PURPOSE: Cleans up a bunch of issues potentially present in requirement statements.
'NOTE: This was written for v9.x and pulled right out of RAM.
    'ACCEPTS:
        'theRS: a requirement statement
        'shouldShallMust: should any instances of should, shall, or must be corrected?
        'ensurePeriodAtEnd: force-adds a period at the end of the requirement statement
        'removeEEG: removes i.e. or e.g. clauses present in parenthesis
        'removeAllParens: removes all parenthetical clauses
    'RETURNS: A requirement statement
Public Function cleanRS(theRS As String, Optional shouldShallMust As Boolean = True, Optional ensurePeriodAtEnd As Boolean = True, Optional removeIEEG As Boolean = False, Optional removeAllParens As Boolean = False) As String
    Dim tmp As String: tmp = theRS
    
    'remove line breaks
    tmp = killLineBreaks(tmp)
    
    'clean up the number presentation
    tmp = Replace(tmp, "three hundred and sixty five (365)", "365")
    tmp = Replace(tmp, "three hundred sixty-five (365)", "365")
    tmp = Replace(tmp, "three hundred and sixty-five (365)", "365")
    tmp = Replace(tmp, "three hundred and sixty- five (365)", "365")
    tmp = Replace(tmp, "three hundred sixty five (365)", "365")
    tmp = Replace(tmp, "three-hundred-sixty-five (365)", "365")
    tmp = Replace(tmp, "365days", "365 days")
    Dim i As Integer
    For i = 1 To 100
        If InStr(1, tmp, "(" & i & ")") > 0 Then
            tmp = Replace(tmp, "(" & i & ")", "")
        End If
    Next i
    tmp = Replace(tmp, "3 year", "three year")
    tmp = Replace(tmp, "fifteen minute", "15 minute")
    tmp = Replace(tmp, "fifty year", "50 year")
    tmp = Replace(tmp, "thirty minute", "30 minute")
    tmp = Replace(tmp, "twenty-four hour", "24 hour")
    tmp = Replace(tmp, "twenty four hour", "24 hour")
    tmp = Replace(tmp, "forty eight hour", "48 hour")
    tmp = Replace(tmp, "forty-eight hour", "48 hour")
    tmp = Replace(tmp, "seventy-two hour", "72 hour")
    tmp = Replace(tmp, "seventy two hour", "72 hour")
    tmp = Replace(tmp, "fourteen character", "14 character")
    tmp = Replace(tmp, "7 day", "seven day")
    tmp = Replace(tmp, "thirty day", "30 day")
    tmp = Replace(tmp, "forty-five day", "45 day")
    tmp = Replace(tmp, "forty five day", "45 day")
    tmp = Replace(tmp, "sixty day", "60 day")
    tmp = Replace(tmp, "ninety day", "90 day")
    tmp = Replace(tmp, "3-year", "three year")
    tmp = Replace(tmp, "fifteen  minute", "15 minute")
    tmp = Replace(tmp, "fifty  year", "50 year")
    tmp = Replace(tmp, "thirty  minute", "30 minute")
    tmp = Replace(tmp, "twenty  four hour", "24 hour")
    tmp = Replace(tmp, "forty  eight hour", "48 hour")
    tmp = Replace(tmp, "seventy  two hour", "72 hour")
    tmp = Replace(tmp, "fourteen  character", "14 character")
    tmp = Replace(tmp, "7 day", "seven day")
    tmp = Replace(tmp, "thirty  day", "30 day")
    tmp = Replace(tmp, "forty  five day", "45 day")
    tmp = Replace(tmp, "sixty  day", "60 day")
    tmp = Replace(tmp, "ninety  day", "90 day")
    
    'clean other weird text and issues from text
    If tmp Like "*third party report*" Then tmp = Replace(tmp, "third party report", "third-party report")
    If tmp Like "*third party provider*" Then tmp = Replace(tmp, "third party provider", "third-party provider")
    If tmp Like "*third party personnel*" Then tmp = Replace(tmp, "third party personnel", "third-party personnel")
    If tmp Like "*third party user*" Then tmp = Replace(tmp, "third party user", "third-party user")
    If tmp Like "*third party support*" Then tmp = Replace(tmp, "third party service", "third-party support")
    If tmp Like "*third party service*" Then tmp = Replace(tmp, "third party service", "third-party service")
    If tmp Like "*third party system*" Then tmp = Replace(tmp, "third party system", "third-party system")
    If tmp Like "*third party contact*" Then tmp = Replace(tmp, "third party contact", "third-party contact")
    If tmp Like "*high risk locations*" Then tmp = Replace(tmp, "high risk locations", "high-risk locations")
    If tmp Like "*decision making roles*" Then tmp = Replace(tmp, "decision making roles", "decision-making roles")
    If tmp Like "*<p>*" Then tmp = Replace(tmp, "<p>", "")
    If tmp Like "*</p>*" Then tmp = Replace(tmp, "</p>", "")
    If tmp Like "*program(s) is (are)*" Then tmp = Replace(tmp, "program(s) is (are)", "programs are")
    If tmp Like "*the organizations*" Then tmp = Replace(tmp, "the organizations", "the organization's")
    If tmp Like "*third-parties*" Then tmp = Replace(tmp, "third-parties", "third parties")
    If tmp Like "*counter-intelligence*" Then tmp = Replace(tmp, "counter-intelligence", "counterintelligence")
    If tmp Like "*personally-owned*" Then tmp = Replace(tmp, "personally-owned", "personally owned")
    If tmp Like "*up-to-date*" Then tmp = Replace(tmp, "up-to-date", "up to date")
    If tmp Like "*rol and*" Then tmp = Replace(tmp, "rol and", "role and")
    If tmp Like "*black list*" Then tmp = Replace(tmp, "black list", "blacklist")
    If tmp Like "*internet*" Then tmp = Replace(tmp, "internet", "Internet")
    If tmp Like "*rol, and*" Then tmp = Replace(tmp, "rol, and", "role, and")
    If tmp Like "*controle *" Then tmp = Replace(tmp, "controle ", "control ")
    If tmp Like "* a updated*" Then tmp = Replace(tmp, " a updated", " an updated")
    If tmp Like "*activies*" Then tmp = Replace(tmp, "activies", "activities")
    If tmp Like "*senor member*" Then tmp = Replace(tmp, "senor member", "senior member")
    If tmp Like "*endored*" Then tmp = Replace(tmp, "endored", "endorsed")
    If tmp Like "*hard-drives*" Then tmp = Replace(tmp, "hard-drives", "hard drives")
    If tmp Like "*Group, shared or generic*" Then tmp = Replace(tmp, "Group, shared or generic", "Group, shared, or generic")
    If tmp Like "*commonly-used*" Then tmp = Replace(tmp, "commonly-used", "commonly used")
    If tmp Like "*cryptographically-protected*" Then tmp = Replace(tmp, "cryptographically-protected", "cryptographically protected")
    If tmp Like "*Visitor and third-party support access is recorded*" Then tmp = Replace(tmp, "Visitor and third-party support access is recorded", "Visitor and third-party support access are recorded")
    
    If shouldShallMust Then
        If tmp Like "*shall be*" Then tmp = Replace(tmp, "shall be", "is")
        If tmp Like "*must document*" Then tmp = Replace(tmp, "must document", "documents")
        If tmp Like "*should be*" Then tmp = Replace(tmp, "should be", "are")
    End If
    
    'remove eg and ie
    If removeIEEG Then
        tmp = removeBracketsAndNumbers(tmp, "(e.g.", ")")
        tmp = removeBracketsAndNumbers(tmp, "[e.g.", "]")
        tmp = removeBracketsAndNumbers(tmp, "(i.e.", ")")
        tmp = removeBracketsAndNumbers(tmp, "(or ", ")")
    ElseIf removeAllParens Then
        tmp = removeBracketsAndNumbers(tmp, "(", ")")
        tmp = removeBracketsAndNumbers(tmp, "[", "]")
    End If
    
    If Right(tmp, 1) = """" Then tmp = Left(tmp, Len(tmp) - 1)
    If ensurePeriodAtEnd Then
        If Not Right(tmp, 1) = "." Then tmp = tmp & "."
        tmp = Trim(tmp)
        If Len(tmp) < 2 Then
            'Debug.Print tmp
        End If
        If Not Right(tmp, 1) = "." Then
            tmp = tmp & "."
        End If
    End If
    
    tmp = Trim(tmp)
    If tmp Like "*" & Chr(160) & "*" Then tmp = Replace(tmp, Chr(160), " ")
    If tmp Like "*  *" Then tmp = Replace(tmp, "  ", " ")
    If tmp Like "* , *" Then tmp = Replace(tmp, " , ", ", ")
    If tmp Like "* ) *" Then tmp = Replace(tmp, " ) ", ") ")
    If tmp Like "* ( *" Then tmp = Replace(tmp, " ( ", " (")
    If tmp Like "* .*" Then tmp = Replace(tmp, " .", ".")
    If tmp Like "* ; *" Then tmp = Replace(tmp, " ; ", "; ")
    If tmp Like "*,.*" Then tmp = Replace(tmp, ",.", ".")
    If tmp Like "*..*" Then tmp = Replace(tmp, "..", ".")
    If tmp Like "*.).*" And Not tmp Like "*etc.)." Then tmp = Replace(tmp, ".).", ").")
    cleanRS = Trim(tmp)
    
End Function

'PURPOSE: Removes bracketed or parenthentical clauses from a string, called by CleanRS
'ACCEPTS:
    '- myString: The string to modify
    '- strStr: the character which marks the start of the bracketed clause to remove
    '- edStr: the character which marks the end of the bracketed clause to remove
'RETURNS:
    '- The modified string
Private Function removeBracketsAndNumbers(myString As String, stStr As String, edStr As String) As String
    Dim tempString As String: tempString = myString
    Dim i As Integer: i = 0
    
    'actually remove the bracketed text
    Do Until InStr(1, tempString, stStr) = 0
        Dim bracketStart As Integer: bracketStart = InStr(tempString, stStr)
        Dim bracketEnd As Integer: bracketEnd = 0
        Do Until bracketEnd > bracketStart Or i > Len(tempString)
            For i = 1 To Len(tempString)
                On Error GoTo oops
                If Mid(tempString, bracketStart + i, 1) = edStr Then
                    bracketEnd = bracketStart + i
                    Exit For
                End If
                On Error GoTo 0
            Next i
        Loop
        tempString = Left(tempString, bracketStart - 1) & Right(tempString, Len(tempString) - bracketEnd)
    Loop
      
    'clean up weird space errors we might have caused
    tempString = Replace(tempString, "  ", " ")
    tempString = Replace(tempString, " , ", ", ")
    tempString = Replace(tempString, " .", ".")
    tempString = Replace(tempString, ",.", ".")
    
    'return the value
    removeBracketsAndNumbers = Trim(tempString)
    
    Exit Function
oops:
    removeBracketsAndNumbers = myString
End Function
