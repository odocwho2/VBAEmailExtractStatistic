Attribute VB_Name = "Main_"

Sub main()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer

Call ExtractEmailInbox
Call ExtractEmailSubFolderLevel1
Call ExtractEmailSubFolderLevel2
Call ExtractEmailSubFolderLevel3
Call ExtractEmailSent
'Fill In Blank Sender
Call DataClean

    Sheets("Extract").Activate
    Cells.Select
    Selection.Copy
    Sheets("SLA").Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("SLA").Activate

Application.StatusBar = "Creating SLA"

Call sortSLA
Call SLA

    Sheets("SLA").Activate
    Cells.Select
    Selection.Copy
    Sheets("SLA NoDuplicates").Select
    Cells.Select
    ActiveSheet.Paste
    Sheets("SLA NoDuplicates").Activate
    
Call SLANoDuplicates
Call Statistics

Application.StatusBar = ""
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Done extracting emails in " & MinutesElapsed
'''''
'#AndreiT 2/11/2021
'''''
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub


Sub ExtractEmailInbox()
k = 0
    Worksheets("Extract").Select
    MailBe = Worksheets("Interface").Cells.Range("B6").Value
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    Dim myFolder As Object  ' Create a folder object.
    Set myFolder = objNSpace.Folders(MailBe).Folders("Inbox")
    Dim objItem As Object
    Dim iRows, iCols As Integer
    iRows = 1
            Cells(iRows, 1) = "Sender"
            Cells(iRows, 2) = "To"
            Cells(iRows, 3) = "Subject"
            Cells(iRows, 4) = "Received"
            Cells(iRows, 5) = "ConversationID"
            Cells(iRows, 6) = "Email Source"
            Cells(iRows, 7) = "SLA"
    Range("A1:G1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    iRows = 2
    StartDate = Worksheets("Interface").Cells.Range("B2").Value
    EndDate = Worksheets("Interface").Cells.Range("B4").Value
    ' Loop through each item in the folder.
    For Each objItem In myFolder.Items
     AllE = myFolder.Items.Count
        If objItem.Class = olMail Then
            Dim objMail As Outlook.MailItem
            Set objMail = objItem
                If objMail.ReceivedTime >= StartDate And objMail.ReceivedTime <= EndDate Then
                    If objMail.SenderEmailType = "SMTP" Then
                    
                        Cells(iRows, 1) = objMail.SenderEmailAddress
                    Else
                        Var = objMail.Sender.Address
                        If Var = "" Then
                         GoTo next_
                        Else
                        
                        End If
                        On Error Resume Next
                        Cells(iRows, 1) = objMail.Sender.GetExchangeUser.PrimarySmtpAddress

                    End If
                    Cells(iRows, 2) = objMail.To
                    Cells(iRows, 3) = objMail.Subject
                    Cells(iRows, 4) = objMail.ReceivedTime
                    Cells(iRows, 5) = objMail.ConversationID
                    Cells(iRows, 6) = myFolder
                Else
                GoTo next_
                End If
        Else
        GoTo next_
        End If

        iRows = iRows + 1
        
next_:
Err.Clear
k = k + 1
Application.StatusBar = "Total emails present in folder level root " & AllE & " current email " & k
    Next
    
    Set objMail = Nothing
   
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing

    
    
End Sub

Sub ExtractEmailSubFolderLevel1()
k = 0
    Worksheets("Extract").Select
     MailBe = Worksheets("Interface").Cells.Range("B6").Value
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    Dim myFolder As Object  ' Create a folder object.
    Set myFolder = objNSpace.Folders(MailBe).Folders("Inbox")
    'Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    Dim SubFolder As Outlook.MAPIFolder
    
    Dim objItem As Object
    Dim iRows, iCols As Integer
    
    lrow = Cells(Rows.Count, 5).End(xlUp).Row

    iRows = lrow + 1
    StartDate = Worksheets("Interface").Cells.Range("B2").Value
    EndDate = Worksheets("Interface").Cells.Range("B4").Value
    ' Loop through each item in the folder.
   For Each SubFolder In myFolder.Folders
    For Each objItem In SubFolder.Items
     AllE = SubFolder.Items.Count
        If objItem.Class = olMail Then
            Dim objMail As Outlook.MailItem
            Set objMail = objItem
                If objMail.ReceivedTime >= StartDate And objMail.ReceivedTime <= EndDate Then
                    If objMail.SenderEmailType = "SMTP" Then
                    
                        Cells(iRows, 1) = objMail.SenderEmailAddress
                    Else
                        Var = objMail.Sender.Address
                        If Var = "" Then
                         GoTo next_
                        Else
                        
                        End If
                        On Error Resume Next
                        Cells(iRows, 1) = objMail.Sender.GetExchangeUser.PrimarySmtpAddress

                    End If
                    Cells(iRows, 2) = objMail.To
                    Cells(iRows, 3) = objMail.Subject
                    Cells(iRows, 4) = objMail.ReceivedTime
                    Cells(iRows, 5) = objMail.ConversationID
                    Cells(iRows, 6) = myFolder & "-" & SubFolder
                Else
                GoTo next_
                End If
        Else
        GoTo next_
        End If

        iRows = iRows + 1
next_:
Err.Clear
k = k + 1
Application.StatusBar = "Total emails present in subfolder level 1 " & AllE & " current email " & k
    Next
    Next
    
    Set objMail = Nothing
   
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing

    
    
End Sub

Sub ExtractEmailSubFolderLevel2()
k = 0
    Worksheets("Extract").Select
     MailBe = Worksheets("Interface").Cells.Range("B6").Value
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    Dim myFolder As Object  ' Create a folder object.
    Set myFolder = objNSpace.Folders(MailBe).Folders("Inbox")
    'Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    Dim SubFolder As Outlook.MAPIFolder
    Dim SubFolderLv2 As Outlook.MAPIFolder
    
    Dim objItem As Object
    Dim iRows, iCols As Integer
    
    lrow = Cells(Rows.Count, 5).End(xlUp).Row

    iRows = lrow + 1
    StartDate = Worksheets("Interface").Cells.Range("B2").Value
    EndDate = Worksheets("Interface").Cells.Range("B4").Value
    ' Loop through each item in the folder.
   For Each SubFolder In myFolder.Folders
    For Each SubFolderLv2 In SubFolder.Folders
        For Each objItem In SubFolderLv2.Items
        AllE = SubFolderLv2.Items.Count
            If objItem.Class = olMail Then
                Dim objMail As Outlook.MailItem
                Set objMail = objItem
                    If objMail.ReceivedTime >= StartDate And objMail.ReceivedTime <= EndDate Then
                        If objMail.SenderEmailType = "SMTP" Then
                        
                            Cells(iRows, 1) = objMail.SenderEmailAddress
                        Else
                            Var = objMail.Sender.Address
                            If Var = "" Then
                             GoTo next_
                            Else
                            
                            End If
                            On Error Resume Next
                            Cells(iRows, 1) = objMail.Sender.GetExchangeUser.PrimarySmtpAddress
    
                        End If
                        Cells(iRows, 2) = objMail.To
                        Cells(iRows, 3) = objMail.Subject
                        Cells(iRows, 4) = objMail.ReceivedTime
                        Cells(iRows, 5) = objMail.ConversationID
                        Cells(iRows, 6) = myFolder & "-" & SubFolder & "-" & SubFolderLv2
                    Else
                    GoTo next_
                    End If
            Else
            GoTo next_
            End If
    
            iRows = iRows + 1
next_:
    Err.Clear
        k = k + 1
        Application.StatusBar = "Total emails present in subfolder level 2 " & AllE & " current email " & k
        Next
    Next
Next
    Set objMail = Nothing
   
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing

    
    
End Sub

Sub ExtractEmailSubFolderLevel3()
k = 0
    Worksheets("Extract").Select
     MailBe = Worksheets("Interface").Cells.Range("B6").Value
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    Dim myFolder As Object  ' Create a folder object.
    Set myFolder = objNSpace.Folders(MailBe).Folders("Inbox")
    'Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    Dim SubFolder As Outlook.MAPIFolder
    Dim SubFolderLv2 As Outlook.MAPIFolder
    
    Dim objItem As Object
    Dim iRows, iCols As Integer
    
    lrow = Cells(Rows.Count, 5).End(xlUp).Row

    iRows = lrow + 1
    StartDate = Worksheets("Interface").Cells.Range("B2").Value
    EndDate = Worksheets("Interface").Cells.Range("B4").Value
    ' Loop through each item in the folder.
   For Each SubFolder In myFolder.Folders
    For Each SubFolderLv2 In SubFolder.Folders
        For Each SubFolderLv3 In SubFolderLv2.Folders
        For Each objItem In SubFolderLv3.Items
            AllE = SubFolderLv3.Items.Count
            If objItem.Class = olMail Then
                Dim objMail As Outlook.MailItem
                Set objMail = objItem
                    If objMail.ReceivedTime >= StartDate And objMail.ReceivedTime <= EndDate Then
                        If objMail.SenderEmailType = "SMTP" Then
                        
                            Cells(iRows, 1) = objMail.SenderEmailAddress
                        Else
                            Var = objMail.Sender.Address
                            If Var = "" Then
                             GoTo next_
                            Else
                            
                            End If
                            On Error Resume Next
                            Cells(iRows, 1) = objMail.Sender.GetExchangeUser.PrimarySmtpAddress
    
                        End If
                        Cells(iRows, 2) = objMail.To
                        Cells(iRows, 3) = objMail.Subject
                        Cells(iRows, 4) = objMail.ReceivedTime
                        Cells(iRows, 5) = objMail.ConversationID
                        Cells(iRows, 6) = myFolder & "-" & SubFolder & "-" & SubFolderLv2 & "-" & SubFolderLv3
                    Else
                    GoTo next_
                    End If
            Else
            GoTo next_
            End If
    
            iRows = iRows + 1
next_:
    Err.Clear
        k = k + 1
        Application.StatusBar = "Total emails present in subfolder level 3 " & AllE & " current email " & k
        Next
    Next
Next
Next
    Set objMail = Nothing
   
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing

    
    
End Sub

Sub ExtractEmailSent()
k = 0
     Worksheets("Extract").Select
     MailBe = Worksheets("Interface").Cells.Range("B6").Value
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    Dim objNSpace As Object     ' Create and Set a NameSpace OBJECT.
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    Dim myFolder As Object  ' Create a folder object.
    Set myFolder = objNSpace.Folders(MailBe).Folders("Sent Items")
    'Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
  
    
    Dim objItem As Object
    Dim iRows, iCols As Integer
    
    lrow = Cells(Rows.Count, 5).End(xlUp).Row

    iRows = lrow + 1
    StartDate = Worksheets("Interface").Cells.Range("B2").Value
    EndDate = Worksheets("Interface").Cells.Range("B4").Value
    ' Loop through each item in the folder.

    For Each objItem In myFolder.Items
        AllE = myFolder.Items.Count
        If objItem.Class = olMail Then
            Dim objMail As Outlook.MailItem
            Set objMail = objItem
                If objMail.SentOn >= StartDate And objMail.SentOn <= EndDate Then
                    If objMail.SenderEmailType = "SMTP" Then
                    
                        Cells(iRows, 1) = objMail.SenderEmailAddress
                    Else
                        Var = objMail.Sender.Address
                        If Var = "" Then
                         GoTo next_
                        Else
                        
                        End If
                        On Error Resume Next
                        Cells(iRows, 1) = objMail.Sender.GetExchangeUser.PrimarySmtpAddress

                    End If
                    Cells(iRows, 2) = objMail.To
                    Cells(iRows, 3) = objMail.Subject
                    Cells(iRows, 4) = objMail.ReceivedTime
                    Cells(iRows, 5) = objMail.ConversationID
                    Cells(iRows, 6) = myFolder
                Else
                GoTo next_
                End If
        Else
        GoTo next_
        End If

        iRows = iRows + 1
next_:
Err.Clear
k = k + 1
Application.StatusBar = "Total emails present in sent " & AllE & " current email " & k
    Next
    
    Set objMail = Nothing
   
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing

    
    
End Sub

Sub SLA()

i = 2
j = 3

lrow = Cells(Rows.Count, 5).End(xlUp).Row

For i = 2 To lrow

    For j = i + 1 To lrow
    
        If Cells.Range("E" & j) <> Cells.Range("E" & i) Then
            Cells.Range("G" & i) = Abs(Cells.Range("D" & i) - Cells.Range("D" & j - 1))
            i = j - 1
            Exit For
            
        Else
        
        End If
    
    Next

Next

End Sub


Sub sortSLA()
lrow = Cells(Rows.Count, 5).End(xlUp).Row

    Range("E1").Select
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
    
    ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "D1:D" & lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "E1:E" & lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SLA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub SLANoDuplicates()


lrow = Cells(Rows.Count, 5).End(xlUp).Row
ActiveSheet.Range("$A$1:$G$" & lrow).AutoFilter Field:=7, Criteria1:="="

ActiveSheet.Range("$A$1:$G$" & lrow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
Selection.AutoFilter
End Sub
Sub DataClean()
   lrow = Cells(Rows.Count, 5).End(xlUp).Row
    Range("A2:A" & lrow).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    
    Selection.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Selection.Replace What:="", Replacement:="NoSender", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

End Sub
Sub Statistics()
UserSent = Worksheets("Interface").Cells.Range("B6").Value

Sheets("SLA NoDuplicates").Select
lrow = Cells(Rows.Count, 5).End(xlUp).Row
    Range("E2:E" & lrow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Statistics").Select
    Range("AI2").Select
    ActiveSheet.Paste
    
Sheets("SLA NoDuplicates").Select
    Range("C2:C" & lrow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Statistics").Select
    Range("AJ2").Select
    ActiveSheet.Paste
 

Sheets("SLA NoDuplicates").Select
    Range("A2:A" & lrow).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Statistics").Select
    Range("AM2").Select
    ActiveSheet.Paste

Range("AK2").Select
Selection.AutoFill Destination:=Range("AK2:AK" & lrow)
Range("AH2").Select
Selection.AutoFill Destination:=Range("AH2:AH" & lrow)
Range("AN2").Select
lrow2 = Cells(Rows.Count, 39).End(xlUp).Row
Selection.AutoFill Destination:=Range("AN2:AN" & lrow2)


Columns("AM:AM").Select
    ActiveSheet.Range("$AM$1:$AM$" & lrow2).RemoveDuplicates Columns:=1, Header:= _
        xlNo
        
lrow2 = Cells(Rows.Count, 39).End(xlUp).Row
'remove sent count
For k = 2 To lrow2

    If Cells.Range("AM" & k) = UserSent Then
        Cells.Range("AM" & k).Clear
        Exit For
        
    Else
    
    End If

Next

End Sub
Sub ClearOldData()
Sheets("Extract").Select
    lrow = Cells(Rows.Count, 5).End(xlUp).Row
    Cells.Range("A2:G" & lrow).Clear
Sheets("SLA NoDuplicates").Select
    lrow = Cells(Rows.Count, 5).End(xlUp).Row
    Cells.Range("A2:G" & lrow).Clear
Sheets("SLA").Select
    lrow = Cells(Rows.Count, 5).End(xlUp).Row
    Cells.Range("A2:G" & lrow).Clear
Sheets("Statistics").Select
lrow = Cells(Rows.Count, 35).End(xlUp).Row
Range("AI2:AI" & lrow).Clear
Range("AJ2:AJ" & lrow).Clear
lrow = Cells(Rows.Count, 39).End(xlUp).Row
Range("AM2:AM" & lrow).Clear
End Sub
