Attribute VB_Name = "WorkingModule"
Option Explicit

Sub PM_Manager()

    Dim wkbManager As Workbook 'Source Code Workbook
    Dim wkbPMData As Workbook 'Pegasus PM data
    Dim bteFeedback As Byte
    Dim strPMdataFilePath As String
    Dim strEmailTemplateFilePath As String
    Dim intYear(1) As Integer 'i=0 current year i=1 next year
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intTotalEntries(1) As Integer
    Dim PMData() As String
    Dim intEntry As Integer
    Dim strSortRange As String
    Dim strWholeRange As String
    Dim NoOfEngineers As Integer
    Dim Engineers() As String
    Dim EngineersEmailAddresses() As String
    Dim NoPerEngineer() As Integer
    Dim EngineerDataSorted() As String
    Dim MaxNoPerEngineer As Integer
    Dim strTextLine() As String
    Dim WordFileContainingEmailBody As String
    Dim WinApp As Object
    Dim EmailBody As String
    Dim EmailSubject As String
    Dim EmailOutline() As String
    Dim FullEmailBody As String
    Dim FaraisMessage As String
    Dim ErrorNo As Integer
    Dim strErrorMessage As String
    Dim AnyLocalPMs As Boolean
    Dim PreviousYear As Integer
    Dim PreviousYearsPMFound As Boolean
    
'On Error GoTo ErrorCatch
ErrorNo = 1

'Ask user to confirm running script
        Set wkbManager = ActiveWorkbook
        bteFeedback = MsgBox("Do you wish to run the automated PM E-mail reminder generator?", vbYesNo, "Automated PM Reminder E-mail")
        If bteFeedback = 7 Then Exit Sub
        
'Grab file paths & open PM data file
        strEmailTemplateFilePath = Range("OutlookTemplate").Value
        strPMdataFilePath = Range("PMdata").Value
        WordFileContainingEmailBody = Range("WordFile").Value
        ErrorNo = 475
        Workbooks(strPMdataFilePath).Activate
        ErrorNo = 1
        Set wkbPMData = ActiveWorkbook
        
'Define current & next years
        intYear(0) = Year(Now())
        intYear(1) = Year(Now()) + 1
        
'Determine no of Reminder Dates to scan
            Sheets(CStr(intYear(0))).Select
            Range("BB1").Value = "=COUNT(A:A)"
            intTotalEntries(0) = Range("BB1").Value
            Range("BC1").Value = "=COUNT(" & Chr(39) & intYear(1) & Chr(39) & "!A:A)"
            intTotalEntries(1) = Range("BC1").Value
            Range("BB1").ClearContents
            Range("BC1").ClearContents
            ReDim PMData(intTotalEntries(0) + intTotalEntries(1), 8) As String
            intEntry = 0
            
'Scan Reminder Dates
        For i = 0 To 1
            Sheets(CStr(intYear(i))).Select
            Call RemoveEmptyRows(intTotalEntries(i))
            
            For intRow = 2 To intTotalEntries(i) + 1
                If IsEmpty(Cells(intRow, 18)) = False Then
                    If IsEmpty(Cells(intRow, 19)) = True Then
                        Range("BB" & intRow).Value = "=R" & intRow & "-NOW()"
                        If Range("BB" & intRow).Value < 1 Then
                            PMData(intEntry, 0) = Cells(intRow, 1).Value 'S/N
                            PMData(intEntry, 1) = Cells(intRow, 2).Value 'IN#
                            PMData(intEntry, 2) = Cells(intRow, 3).Value 'Company Name
                            PMData(intEntry, 3) = Cells(intRow, 7).Value 'Customer ID
                            PMData(intEntry, 4) = Cells(intRow, 9).Value 'Unit
                            PMData(intEntry, 5) = Cells(intRow, 22).Value 'PM Due / CONTRACT EXPIRES
                            PMData(intEntry, 6) = CStr(intYear(i)) 'Sheet PM was found on
                            PMData(intEntry, 7) = Cells(intRow, 20).Value 'Engineer
                            PMData(intEntry, 8) = Cells(intRow, 21).Value 'E-mail
                            
                            
                            Range("S" & intRow).Value = Date
                            intEntry = intEntry + 1
                        End If
                    End If
                End If
            Next intRow
            Columns("BB:BB").Select
            Selection.ClearContents
            Range("A1").Select
        Next i
        
'Check if there are no entries
        If intEntry = 0 Then
            GoTo NoPendingPMs
        End If
        
'Drop Data Into Data Organization worksheet
        intEntry = intEntry - 1
        wkbManager.Activate
        Sheets("Organize Data").Select
        For i = 0 To intEntry
            intRow = 1 + i
            For j = 0 To 8
                intColumn = 1 + j
                Cells(intRow, intColumn).Value = PMData(i, j)
            Next j
            'Cells(intRow, 26).Value = PMData(i, j)
        Next i
        
'Sort by engineer
        strSortRange = "H1:H" & intRow
        strWholeRange = "A1:I" & intRow
        Range(Cells(1, 1), Cells(intRow, intColumn)).Select 'Fails here when there are 0 entries
        ActiveWorkbook.Worksheets("Organize Data").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Organize Data").Sort.SortFields.Add Key:=Range( _
            strSortRange), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Organize Data").Sort
            .SetRange Range(strWholeRange)
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
'Determine how many engineers and who they are
        strWholeRange = "K1:L" & intRow
        strSortRange = "H1:I" & intRow
        Range(strSortRange).Select
        Selection.Copy
        Range(strWholeRange).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Application.DisplayAlerts = False
        ActiveSheet.Range(strWholeRange).RemoveDuplicates Columns:=2, Header:=xlNo
        Range("M1").Value = "=COUNTA(K:K)-1"
        NoOfEngineers = Range("M1").Value
        ReDim Engineers(NoOfEngineers) As String
        ReDim EngineersEmailAddresses(NoOfEngineers) As String
        ReDim NoPerEngineer(NoOfEngineers) As Integer
        For i = 0 To NoOfEngineers
            intRow = i + 1
            Engineers(i) = Range("K" & intRow).Value
            EngineersEmailAddresses(i) = Range("L" & intRow).Value
            Range("J" & intRow).Value = "=COUNTIF(H:H,K" & intRow & ")"
            NoPerEngineer(i) = Range("J" & intRow).Value - 1
        Next i
        Range("M2").Value = "=MAX(J:J)"
        MaxNoPerEngineer = Range("M2").Value - 1
        
'Determine if there are any local PM's
        AnyLocalPMs = FindString("00_Local")
        
'Populate engineer data array
        ReDim EngineerDataSorted(NoOfEngineers, MaxNoPerEngineer, 7) As String
        ReDim strTextLine(NoOfEngineers, MaxNoPerEngineer) As String
        ReDim EmailOutline(NoOfEngineers) As String
        intRow = 0
        For i = 0 To NoOfEngineers
            For j = 0 To NoPerEngineer(i)
                intRow = intRow + 1
                intColumn = 0
                For k = 0 To 6
                    intColumn = intColumn + 1
                    EngineerDataSorted(i, j, k) = Cells(intRow, intColumn).Value
                    If k < 6 Then
                        Call DropEngineerDataToWorksheet("Organize Data", Engineers(i), k, EngineerDataSorted(i, j, k))
                    End If
                Next k
            Next j
        Next i
        Cells.Select
        Selection.ClearContents
        
'Determine date of last PM & populate into engineer data array
        wkbPMData.Activate
        
        For i = 0 To NoOfEngineers
            For j = 0 To NoPerEngineer(i)
                EngineerDataSorted(i, j, 7) = EngineerDataSorted(i, j, 6)
                PreviousYear = CInt(EngineerDataSorted(i, j, 7)) - 1
                Sheets(CStr(PreviousYear)).Select
                
                '(InstanceNumber, SerialNumber, PreviousPMDate)
                PreviousYearsPMFound = FindPreviousPMEntry(EngineerDataSorted(i, j, 1), EngineerDataSorted(i, j, 0), EngineerDataSorted(i, j, 6))
                
                If PreviousYearsPMFound = False Then
                    PreviousYear = CInt(EngineerDataSorted(i, j, 7)) - 2
                    Sheets(CStr(PreviousYear)).Select
                    '(InstanceNumber, SerialNumber, PreviousPMDate)
                    PreviousYearsPMFound = FindPreviousPMEntry(EngineerDataSorted(i, j, 1), EngineerDataSorted(i, j, 0), EngineerDataSorted(i, j, 6))
                End If
                
                wkbManager.Activate
                Call DropEngineerDataToWorksheet("Sheet1", Engineers(i), 6, EngineerDataSorted(i, j, 6))
                wkbPMData.Activate
                
            Next j
        Next i
        
        wkbManager.Activate
        
'Generate e-mails and send
        
        If AnyLocalPMs = True Then
            Engineers(NoOfEngineers) = "Local PM's"
        End If
        
        Set WinApp = GetObject(WordFileContainingEmailBody)
        EmailBody = WinApp.Range(Start:=WinApp.Paragraphs(1).Range.Start, End:=WinApp.Paragraphs(WinApp.Paragraphs.Count).Range.End)
        Set WinApp = Nothing
        j = 0
        k = 0
        For i = 0 To NoOfEngineers
            For j = 0 To NoPerEngineer(i)
                strTextLine(i, j) = "     - S/N: "
                For k = 0 To 6
                    strTextLine(i, j) = strTextLine(i, j) & EngineerDataSorted(i, j, k)
                    Select Case k
                        Case 0
                            strTextLine(i, j) = strTextLine(i, j) & "   IN#: "
                        Case 1
                            strTextLine(i, j) = strTextLine(i, j) & "   Cust. Name: "
                        Case 2
                            strTextLine(i, j) = strTextLine(i, j) & "   Cust. ID: "
                        Case 3
                            strTextLine(i, j) = strTextLine(i, j) & Chr(13) & "          Unit: "
                        Case 4
                            strTextLine(i, j) = strTextLine(i, j) & "   PM Due: "
                        Case 5
                            strTextLine(i, j) = strTextLine(i, j) & "   Last PM: "
                        Case 6
                            strTextLine(i, j) = strTextLine(i, j) & Chr(13) & Chr(13)
                    End Select
                Next k
                EmailOutline(i) = EmailOutline(i) & strTextLine(i, j)
            Next j
            FullEmailBody = EmailOutline(i) & Chr(13)
            
            If AnyLocalPMs = False Then
                Call CreateAndSendEmail(EngineersEmailAddresses(i), Engineers(i), strEmailTemplateFilePath, EmailBody & Chr(13) & FullEmailBody)
            ElseIf AnyLocalPMs = True And i < NoOfEngineers Then
                Call CreateAndSendEmail(EngineersEmailAddresses(i), Engineers(i), strEmailTemplateFilePath, EmailBody & Chr(13) & FullEmailBody)
            End If
            
            FaraisMessage = FaraisMessage & Engineers(i) & Chr(13) & FullEmailBody & Chr(13)
        Next i
        
        Call CreateAndSendFaraisEmail("farai_rukunda@leco.com", "Farai", "PM Reminders - " & Date & Chr(13) & Chr(13) & FaraisMessage, "jennifer_fry@leco.com")
        Sheets(1).Select
        Range("A1").Select
        
        Exit Sub
        
        
NoPendingPMs:
        MsgBox "No Pending PM reminders", vbInformation, "PM-Manager"
Exit Sub


ErrorCatch:
    strErrorMessage = "Undefined Error"
    Select Case ErrorNo
        Case 475
            strErrorMessage = strPMdataFilePath & " - Source file not opened."
    End Select
    
    MsgBox strErrorMessage, vbCritical, "Error #: " & ErrorNo

End Sub

Private Sub RemoveEmptyRows(ByVal NoOfEntries As Integer)
Dim bolContinue As Boolean
Dim Row As Integer
Const Column As Byte = 1
Dim i As Integer

    bolContinue = True
    Row = 2
    i = 0
    
    Do While bolContinue = True
        Cells(Row, Column).Select
        If IsEmpty(Cells(Row, Column)) <> False Then
            Rows(Row & ":" & Row).Select
            Selection.Delete Shift:=xlUp
            Row = Row - 1
        ElseIf IsEmpty(Cells(Row, Column)) = False Then
            i = i + 1
        End If
        
        If i = NoOfEntries Then
            bolContinue = False
        End If
        Row = Row + 1
    Loop

End Sub

Private Sub CreateAndSendEmail(ByVal EMailAddress As String, ByVal ContactName As String, ByVal OutlookTemplateFilePath As String, ByVal EmailBody As String)
    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(OutlookTemplateFilePath)
    With OutMail
        .To = EMailAddress
        .Body = "Hello " & ContactName & "," & vbNewLine & EmailBody
        .Display
        '.Send
    End With
End Sub

Private Sub CreateAndSendFaraisEmail(ByVal EMailAddress As String, ByVal ContactName As String, ByVal EmailBody As String, ByVal CC_EmailAddress As String)
    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(olMailItem)
    With OutMail
        .To = EMailAddress
        .Body = EmailBody
        .Display
        .Importance = olImportanceHigh
        .Subject = "PM Notifications: " & Date
        .CC = CC_EmailAddress
        '.Send
    End With
End Sub

Private Sub DropEngineerDataToWorksheet(ByVal strSheet1 As String, ByVal strSheet2 As String, ByVal FieldNo As Integer, ByVal strEngineerValue As String)

    Dim Row As Integer
    Dim Column As Integer
    
    Sheets(strSheet2).Select
    Column = 4 + FieldNo
    
    Cells(2, Column).Select
    Selection.End(xlDown).Select
    Row = ActiveCell.Row + 1
    
    Cells(Row, Column).Value = strEngineerValue
    Cells(Row, 3).Value = Date
    Sheets(strSheet1).Select
    
    
End Sub

Private Function FindString(ByVal StringToFind As String) As Boolean
On Error GoTo StringNotFound

    Cells.Find(what:=StringToFind, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
FindString = True
Exit Function
StringNotFound:
FindString = False
End Function

Private Function FindPreviousPMEntry(ByVal InstanceNumber As String, ByVal SerialNumber As String, PreviousPMDate As String) As Boolean
Dim ContinueStep_1 As Boolean
Dim ContinueStep_2 As Boolean
Dim FoundPMDate As String
Dim PM_Row As String
Const SerialNumberColumn As Integer = 1
Const LastPMDateColumn As Integer = 12

    
    Cells(1, 1).Select
    ContinueStep_1 = FindString(InstanceNumber)
    
    If ContinueStep_1 = True Then
        PM_Row = ActiveCell.Row
        
    ElseIf ContinueStep_1 = False Then
        FindPreviousPMEntry = False
        PreviousPMDate = "N/A"
        Exit Function
    End If
    
    If Cells(PM_Row, SerialNumberColumn).Value = SerialNumber Then
        ContinueStep_2 = True
    Else
        ContinueStep_2 = False
    End If
    
    If ContinueStep_2 = True Then
        FoundPMDate = Cells(PM_Row, LastPMDateColumn).Value
    ElseIf ContinueStep_2 = False Then
        FindPreviousPMEntry = False
        PreviousPMDate = "N/A"
        Exit Function
    End If
    
    If FoundPMDate = "" Then
        FindPreviousPMEntry = False
        PreviousPMDate = "N/A"
        Exit Function
    Else
        FindPreviousPMEntry = True
        PreviousPMDate = FoundPMDate
        Exit Function
    End If
    

End Function
