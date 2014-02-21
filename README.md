vba
===

some vba work
Public All_Names()
Dim ciscoid As String
Dim ciscopw As String
Dim firstcc As String
Dim secondcc As String
Dim thirdcc As String
Dim firstname As String
Dim lastname As String
Dim CECarea As String
Dim masterfilename As String
Dim masterfilepath As String
Dim masterfileprename As String


Sub createsheetnames() ' this sub creates an array 2 columns, 376 rows to store sheet names
Sheets("Hierarchy").Activate
Height = Range(Cells(5, 7), Selection.End(xlDown)).Count
Dim x As Long
Dim nameofsheet As String
Dim numberofsheets As Long
Dim placeholder As Long

ReDim All_Names(Worksheets.Count, 9) 'there are 9 fields
masterfilename = Range("N1").Value
masterfilepath = Range("K1").Value
masterfileprename = Range("P1").Value

'For x = 4 To 126
'Debug.Print Worksheets(x).Name
'Debug.Print Worksheets(x + 124).Name
'Debug.Print Worksheets(x + 248).Name
'Next
'Debug.Print Cells(5, 7).Value

'Dim trial As String
'trial = "Americas"
'Cells(1, 1).Value = Workbooks("Q2'12 OD Governance Package - Americas.xlsm").Sheets(trial).Cells(7, 7).Value

For x = 0 To 122 'Worksheets.Count - 5
  All_Names(x, 1) = Cells(x + 5, 6).Value ' number of sheets in the file
  All_Names(x, 0) = Cells(x + 5, 4).Value ' name of Area / file
  All_Names(x, 2) = Cells(x + 5, 8).Value ' OD Cisco ID
  All_Names(x, 3) = Cells(x + 5, 9).Value 'password
  All_Names(x, 4) = Cells(x + 5, 10).Value '1st CC
  All_Names(x, 5) = Cells(x + 5, 11).Value '2nd CC
  All_Names(x, 6) = Cells(x + 5, 12).Value '3rd CC
  All_Names(x, 7) = Cells(x + 5, 13).Value 'First Name
  All_Names(x, 8) = Cells(x + 5, 14).Value 'Last Name
  All_Names(x, 9) = Cells(x + 5, 15).Value 'CEC Area
  
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 0)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 4)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 5)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 6)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 7)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 8)
'Debug.Print All_Names(x, 0) & " " & All_Names(x, 9)

Next
   
For x = 0 To Height - 3
   nameofsheet = All_Names(x, 0)
   numberofsheets = All_Names(x, 1) 'second column in array is numbers: contrary to Hierarchy sheet
   placeholder = x
   ciscoid = All_Names(x, 2)
   ciscopw = All_Names(x, 3)
  
  
 If All_Names(x, 4) = "" Then
firstcc = ""
Else
firstcc = All_Names(x, 4) & "@cisco.com;"
End If
  
If All_Names(x, 5) = "" Then
secondcc = ""
Else
secondcc = All_Names(x, 5) & "@cisco.com;"
End If

If All_Names(x, 6) = "" Then
thirdcc = ""
Else
thirdcc = All_Names(x, 6) & "@cisco.com;"
End If

  firstname = All_Names(x, 7)
  lastname = All_Names(x, 8)
  
  areaname = All_Names(x, 9) 'this CEC Area name is for the person, not the File name. For Example, Conor's area would be his CEC title (ECC Area)
   
    
copysheetsover nameofsheet, numberofsheets, placeholder
'Debug.Print nameofsheet & " " & numberofsheets&; " " & placeholder
'Debug.Print All_Names(placeholder + 1 + x, 1)
Next


End Sub

Sub copysheetsover(FileName As String, filesheets As Long, sheetplaceholder As Long)
Dim Mystr As String
Application.ScreenUpdating = False
Dim firstsheet As String
Dim secondsheet As String
Dim thirdsheet As String

Dim activeworkbookname As String
Mystr = FileName & " Dashboard"
Sheets(Mystr).Copy
  


Workbooks(Workbooks.Count).SaveAs _
    FileName:=masterfilepath & _
              masterfileprename & FileName & ".xlsx", CreateBackup:=False 'Password:=ciscopw


  
  If filesheets > 1 Then
  
  For x = 0 To filesheets - 2  'brings in Dashboards
    firstsheet = All_Names(x + 1 + sheetplaceholder, 0) & " Dashboard"
         Workbooks(masterfilename).Sheets(firstsheet).Copy _
        after:=ActiveWorkbook.Worksheets(Worksheets.Count)
  Next
          
   For x = 0 To filesheets - 1 ' brings in Trends
    secondsheet = All_Names(x + sheetplaceholder, 0) & " Trend"
         Workbooks(masterfilename).Sheets(secondsheet).Copy _
        after:=ActiveWorkbook.Worksheets(Worksheets.Count)
  Next
  
    For x = 0 To filesheets - 1  ' brings in Top Discounts
    thirdsheet = All_Names(x + sheetplaceholder, 0) & " Top Dsct"
          Workbooks(masterfilename).Sheets(thirdsheet).Copy _
        after:=ActiveWorkbook.Worksheets(Worksheets.Count)
  Next

activeworkbookname = ActiveWorkbook.Name


TOC sheetplaceholder, filesheets
Workbooks(activeworkbookname).Activate
Debug.Print ActiveWorkbook.Name
        
        Workbooks(masterfilename).Sheets("FAQ's").Copy _
        before:=ActiveWorkbook.Worksheets(1)

        Workbooks(masterfilename).Sheets("Table of Contents").Copy _
        before:=ActiveWorkbook.Worksheets(2)

        
ActiveWorkbook.Save


 ' emailbody FileName, ciscoid   'send workbook as email


ActiveWorkbook.Close

Else

  'brings in Dashboards
    ' firstsheet = All_Names(sheetplaceholder, 0) & " Dashboard"
         ' Workbooks("Q2'12 OD Governance Package - Americas.xlsm").Sheets(firstsheet).Copy _
      ' after:=ActiveWorkbook.Worksheets(Worksheets.Count)
  
    ' brings in Trends
     secondsheet = All_Names(sheetplaceholder, 0) & " Trend"
         Workbooks(masterfilename).Sheets(secondsheet).Copy _
        after:=ActiveWorkbook.Worksheets(Worksheets.Count)
' brings in Top Discounts
    thirdsheet = All_Names(sheetplaceholder, 0) & " Top Dsct"
          Workbooks(masterfilename).Sheets(thirdsheet).Copy _
        after:=ActiveWorkbook.Worksheets(Worksheets.Count)
                
        Workbooks(masterfilename).Sheets("FAQ's").Copy _
        before:=ActiveWorkbook.Worksheets(1)
        ActiveWorkbook.Sheets(2).Activate
        
Debug.Print ActiveWorkbook.Name
        
ActiveWorkbook.Save

 ' emailbody FileName, ciscoid  'send workbook as email

ActiveWorkbook.Close

End If


       

End Sub

Sub TOC(rowholder As Long, numofsheets As Long)
Workbooks(masterfilename).Activate

Application.ScreenUpdating = False

    Sheets("Table of Contents").Activate
    ActiveSheet.Range("C4:H600").Select
    Selection.ClearContents

Sheets("TOC").Range("C2").Copy
Sheets("Table of Contents").Range("C3").Select
ActiveSheet.Paste

Sheets("TOC").Activate ' paste Dashboard
ActiveSheet.Range(Cells(5 + rowholder, 4), Cells(5 + rowholder + numofsheets - 1, 8)).Copy

Sheets("Table of Contents").Activate
Cells(4, 4).Select
ActiveSheet.Paste
Sheets("Table of Contents").Range("C2").Copy
Cells(3, 3).Select

Sheets("TOC").Range("C129").Copy
Sheets("Table of Contents").Range("D5").Offset(numofsheets, -1).Select
ActiveSheet.Paste

Sheets("TOC").Activate ' Paste Trend
ActiveSheet.Range(Cells(5 + rowholder + 127, 4), Cells(5 + rowholder + 127 + numofsheets - 1, 8)).Copy
Sheets("Table of Contents").Activate
Sheets("Table of Contents").Cells(5 + numofsheets + 1, 4).Select
ActiveSheet.Paste


Sheets("TOC").Range("C256").Copy
Sheets("Table of Contents").Range("D5").Offset(numofsheets * 2 + 2, -1).Select
ActiveSheet.Paste

Sheets("TOC").Activate 'paste Top Discount
ActiveSheet.Range(Cells(5 + rowholder + 254, 4), Cells(5 + rowholder + 254 + numofsheets - 1, 8)).Copy
Sheets("Table of Contents").Activate
Sheets("Table of Contents").Range("D5").Offset(numofsheets * 2 + 3, 0).Select
ActiveSheet.Paste
ActiveSheet.Range("A1").Activate

End Sub

Sub printnames()

On Error GoTo Err_Handler



Sheets("Hierarchy").Range("AA4").Select
For x = 0 To Worksheets.Count

ActiveCell.Value = Sheets(x + 5).Name

If x = 123 Or x = 247 Then
m = 4
Else: m = 1
End If

Debug.Print x
ActiveCell.Offset(m, 0).Select
Next


Exit_This_Sub:

Exit Sub

Err_Handler:
Resume Exit_This_Sub




End Sub



Sub emailbody(areaname As String, emailid As String)
'Working in Office 2000-2010
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    If ActiveWorkbook.Path <> "" Then
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)

        strbody = "<font size=""3"" face=""Calibri"">" & _
                   "Update: Apologies but we've just discovered that there is a compatibility issue between Excel 2007 and 2010 with regards to passwords.  If you have Excel 2010, the file that was sent to you yesterday should work using the password logic below.   If you have Excel 2007, please disregard yesterday's file and use this file instead, same password logic.  Thanks." & _
                   "<br><br>Hi " & firstname & ",<br>" & _
                   "<br>In the continuous effort to improve the Americas Field Empowerment governance process, the Americas CF Ops team is now distributing OD Packages systematically through Outlook.&nbsp; Attached, you will find your Q2&rsquo;12 OD Package for your area, " & areaname & " Area. As in prior periods, please review your business trends and engage with your in-theater CF and SSF teams to understand and address any potential issues.&nbsp; <br>" & _
                   "<br>Due to sensitive information contained within, we have password protected the file.&nbsp; Your password is as follows:&nbsp; &ldquo;usernamePPYY&rdquo;, where PP is the fiscal period and YY is fiscal year.&nbsp; For example, for John Chambers, his password would be &ldquo;chambers0612&rdquo;.<br>&nbsp;<br>For any questions, please do not hesitate to contact your local CF team or the Americas CF Ops team listed below:&nbsp;<br>" & _
                   "<br>FED Manager - Sean Liu (<a href=""mailto:seliu@cisco.com"">seliu@cisco.com</a>)<br>CF Executive - Jair Hernandez (<a href=""mailto:eliohern@cisco.com"">eliohern@cisco.com</a>)<br>CF Executive - Nick Pecchenino (<a href=""mailto:npecchen@cisco.com"">npecchen@cisco.com</a>) <br><br>Thanks.<br><br>Americas CF Ops<br>" & _
                   "<br><font color = #990000>***This email contains sensitive material and <strong><u>MUST NOT BE</u></strong> shared or distributed*** <br>"

                 ' "Hi,</br><br>&nbsp;</br><br>You are receiving this test email because you have been identified by CF as being an OD for your sales area, " & _
                  'areaname & _
                  '" Area. The Americas CF Ops team will now begin distributing the OD Packages systematically via email and we need to know if you are no longer responsible for this area.&nbsp; The reports will be available next week so please respond by end of day tomorrow, Friday 2/3, to help us ensure the information gets to the proper contact.&nbsp; We appreciate your cooperation. &nbsp;&nbsp;</br><br>   &nbsp;</br><br>   Thanks. &nbsp;&nbsp;</br><br>&nbsp;</br><br>    Americas CF Ops</br>"
                  '& _
                  '"<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:454px;"" width=""454""><tbody><tr><td colspan=""2"">&nbsp;</td></tr><tr><td><p><strong>E. Jair Hernandez</strong><br /><strong>Commercial Finance Solutions Executive</strong></p><p>Commercial Segment<br /><a href=""mailto:eliohern@cisco.com"">eliohern@cisco.com</a><br />Phone: <strong>408.424.6574</strong><br />Mobile: <strong>832.788.6187</strong></p></td><td>&nbsp;</td></tr><tr><td nowrap=""nowrap""><p>"
            '"Hello,<br><br>" & _
                  "Quarter end OD Report has been created for your review.<br><B>" & _
                  ActiveWorkbook.Name & "</B> is attached to this email<br>" & _
                  "Your password to this file is your Cisco User ID and the month & year in the following format usernameMMYY." & _
                  "<br><br>Regards," & _
                  "<br><br> <font color=#990000>***This email contains sensitive material and" & "<B> MUST NOT BE</b>" & " shared or distributed***</font>" & _
                  "<br><br>Commercial Finance Team" & _

        On Error Resume Next
        With OutMail
            .To = emailid & "@cisco.com"
            .cc = firstcc & secondcc & thirdcc
            .BCC = ""
            .Subject = "UPDATE: Q2'12 " & areaname & " Area OD Package"
            .Attachments.Add ActiveWorkbook.FullName
            .HTMLBody = strbody
            .display  'or use .Send/Display
            .Close olpromptforsave
        End With
        On Error GoTo 0

        Set OutMail = Nothing
        Set OutApp = Nothing
   Else
        MsgBox "The ActiveWorkbook does not have a path, Save the file first."
    End If
End Sub



Sub Send_Selection_Or_ActiveSheet_with_MailEnvelope()
    Dim Sendrng As Range

   ' On Error GoTo StopMacro

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
   
   ActiveSheet.Range("k10:s31").Select

    'Note: if the selection is one cell it will send the whole worksheet
    Set Sendrng = ActiveSheet.Range("k10:s31")

    'Create the mail and send it
    With Sendrng

        ActiveWorkbook.EnvelopeVisible = True
        With .Parent.MailEnvelope

            ' Set the optional introduction field thats adds
            ' some header text to the email body.
            .Introduction = "This is a test mail."

            ' In the "With .Item" part you can add more options
            ' See the tips on this Outlook example page.
            ' http://www.rondebruin.nl/mail/tips2.htm
            With .Item
                .To = "eliohern@cisco.com"
                .cc = "npecchen@cisco.com"
                .Subject = "My subject"
                '.send
                .display
            End With

        End With
    End With

'StopMacro:
    'With Application
       ' .ScreenUpdating = True
    '    .EnableEvents = True
  '  End With
'   ActiveWorkbook.EnvelopeVisible = False

End Sub






