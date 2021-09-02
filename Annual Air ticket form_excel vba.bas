Attribute VB_Name = "Module1"
Option Explicit

Public ERP_No As Range
Public Mobile_UAE As Range
Public Email_Id As Range
Public Date_of_Travel As Range
Public Date_of_Return As Range
Public Boarding_Port As Range
Public Boarding_destination As Range
Public DOT_Time As Range
Public DOR_Time As Range
Public Mobile_Home As Range
Public Booking_Ref As Range
Public LPO_No As Range
Public Carrier_Name As Range
Public EWF_ref As Range
Public Fare1  As Range
Public Fare2  As Range
Public First_Name As Range
Public Last_Name As Range
Public op1_ref As Range
Public op1_1 As Range
Public op1_Fare As Range
Public op2_ref As Range
Public op2_1 As Range
Public op2_Fare  As Range
Public Job_Desc As Range
Public Pay_Grade As Range
Public Pay_Start_Date As Range
Public Nationality As Range
Public Comp_Name As Range
Public Dept_Name As Range
Public DOB As Range
Public BU_No As Range
Public BU_Name As Range
Public PP_No As Range
Public PP_Exp_Date As Range
Public Ticket_Destination As Range
Public Status As Range

Sub OptionButton4_Click()
    Dim book_ref_1 As Range
    Set Booking_Ref = Sheet1.Range("G16")
    Set book_ref_1 = Sheet1.Range("C19")
    'Set Fare2 = Sheet1.Range("F21")
Booking_Ref.Value = book_ref_1.Value
End Sub

Sub OptionButton6_Click()
    Dim book_ref_2 As Range
    Set Booking_Ref = Sheet1.Range("G16")
    'Set Fare1 = Sheet1.Range("C21")
    Set book_ref_2 = Sheet1.Range("F19")
Booking_Ref.Value = book_ref_2.Value
End Sub

Sub Add_record()

    Dim lastrow As Long, ws As Worksheet, rngERP As Range, rgFound As Range
    
        Set ws = Sheet4
        Set rgFound = ws.Range("V:X").Find(Sheet1.Range("C5").Value)
        lastrow = ws.Range("V" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    
    'If Sheet1.Range("C5").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet3.Range("A:C"), 1, False) Then
            If IsEmpty(Sheet1.Range("E14").Value) Or IsEmpty(Sheet1.Range("G14").Value) Then
                MsgBox "Please enter valid Mobile Number and Email id.", vbCritical, "Air Ticket Form"
            Exit Sub
            End If
    
    If rgFound Is Nothing Then
    'Sheet1.Range("C5").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet3.Range("A:C"), 1, False) Then
        ws.Range("V" & lastrow).Value = Sheet1.Range("C5").Value 'Adds the TextBox3 into Col A & Last Blank Row
        ws.Range("W" & lastrow).Value = Sheet1.Range("E14").Value  'Adds the ComboBox1 into Col B & Last Blank Row
        ws.Range("X" & lastrow).Value = Sheet1.Range("G14").Value 'Adds the ComboBox1 into Col C & Last Blank Row
        
        MsgBox "New record updated", , "Air Ticket Form"
    Else
        'Sheet1.Range("C5").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet3.Range("A:C"), 1, False)
        'MsgBox "ERP Number already captured.", , "Air Ticket Form"
        Exit Sub
        
    End If
End Sub

Sub Change_ERP()
    'Dim ERP_No As Range, Mobile_No As Range, Email_Id As Range
    Set ERP_No = Sheet1.Range("C5")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set Booking_Ref = Sheet1.Range("G16")
    Set LPO_No = Sheet1.Range("C36")
    Set Carrier_Name = Sheet1.Range("C37")
    Set EWF_ref = Sheet1.Range("G4")
    Set Fare1 = Sheet1.Range("C21")
    Set Fare2 = Sheet1.Range("F21")
    'below are employee details
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    Set Job_Desc = Sheet1.Range("C6")
    Set Pay_Grade = Sheet1.Range("E6")
    Set Pay_Start_Date = Sheet1.Range("G6")
    Set Nationality = Sheet1.Range("C7")
    Set Comp_Name = Sheet1.Range("E7")
    Set Dept_Name = Sheet1.Range("G7")
    Set DOB = Sheet1.Range("C8")
    Set BU_No = Sheet1.Range("E8")
    Set BU_Name = Sheet1.Range("G8")
    Set PP_No = Sheet1.Range("C9")
    Set PP_Exp_Date = Sheet1.Range("E9")
    Set Ticket_Destination = Sheet1.Range("G9")
    Set Status = Sheet1.Range("C4")
    
    Date_of_Travel.ClearContents
    Date_of_Return.ClearContents
    Boarding_Port.ClearContents
    Boarding_destination.ClearContents
    DOT_Time.ClearContents
    DOR_Time.ClearContents
    Mobile_Home.ClearContents
    Mobile_UAE.ClearContents
    Email_Id.ClearContents
    Booking_Ref.ClearContents
    LPO_No.ClearContents
    Carrier_Name.ClearContents
    EWF_ref.ClearContents
    Fare1.ClearContents
    Fare2.ClearContents
    Sheet1.Range("C19").ClearContents
    Sheet1.Range("F19").ClearContents
    Sheet1.Range("B20:D20").ClearContents
    Sheet1.Range("E20:G20").ClearContents
    Sheet1.Range("C29").ClearContents
    Sheet1.Range("C30").ClearContents
    First_Name.ClearContents
    Last_Name.ClearContents
    Job_Desc.ClearContents
    Pay_Grade.ClearContents
    Pay_Start_Date.ClearContents
    Nationality.ClearContents
    Comp_Name.ClearContents
    Dept_Name.ClearContents
    DOB.ClearContents
    BU_No.ClearContents
    BU_Name.ClearContents
    PP_No.ClearContents
    PP_Exp_Date.ClearContents
    Ticket_Destination.ClearContents
    Status.ClearContents
            
        If IsError(Application.Match(ERP_No.Value, Sheet2.Range("B:B"), 0)) Then
            MsgBox "ERP_No not found in Master File. Please contact HR.", vbCritical, "Air Ticket Form"
        Else
            Call ERP_Emp_Lookup
        End If
            
        If IsError(Application.Match(ERP_No.Value, Sheet4.Range("V:V"), 0)) Then
            Mobile_UAE.ClearContents
            Email_Id.ClearContents
        Else
            Call ERP_Lookup
        End If

End Sub

Sub ERP_Lookup()
    If Sheet1.Range("C5").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet4.Range("V:X"), 1, False) Then
        Sheet1.Range("E14").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet4.Range("V:X"), 2, False)
        Sheet1.Range("G14").Value = Application.VLookup(Sheet1.Range("C5").Value, Sheet4.Range("V:X"), 3, False)
    Else
        'do nothing
    End If
End Sub

Sub Update_record()
    Dim RowNo As Integer
    Dim ans As Integer
    
    Set ERP_No = Sheet1.Range("C5")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    
    RowNo = Application.Match(ERP_No.Value, Sheet4.Range("V:V"), 0)
    
        If ERP_No.Value = Application.VLookup(ERP_No.Value, Sheet4.Range("V:X"), 1, False) Then
            If IsEmpty(Mobile_UAE.Value) Or IsEmpty(Email_Id.Value) Then
                MsgBox "Please enter valid Mobile Number and Email id.", vbCritical, "Air Ticket Form"
            Exit Sub
            End If
                    Sheet4.Range("W" & RowNo) = Mobile_UAE.Value
                    Sheet4.Range("X" & RowNo) = Email_Id.Value
                    
                    MsgBox "Contact Details record updated successfully.", , "Air Ticket Form"
            
        Else
            'do nothing
        End If
End Sub

Sub Capture_Contact()
    Dim ans As Integer, rgFound As Range, FoundCell As Range, ERP_in_Master As Range
    Dim ws2 As Worksheet, ws1 As Worksheet, ws4 As Worksheet
    
    Set ws1 = Sheet1
    Set ws2 = Sheet2
    Set ws4 = Sheet4
    Set ERP_No = Sheet1.Range("C5")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set First_Name = Sheet1.Range("E5")
    
    Set FoundCell = ws4.Range("V:V").Find(what:=ERP_No.Value)
    Set rgFound = ws4.Range("V:X").Find(ERP_No.Value)
    Set ERP_in_Master = ws2.Range("B:B").Find(ERP_No.Value)

    
    If rgFound Is Nothing Then
        ans = MsgBox("New Mobile & Email_id found for " & First_Name & " Do you want to update?", vbYesNo + vbQuestion, "For you information")
        If ans = vbYes Then
            Call Add_record
        End If
    Else
        If Mobile_UAE.Value = ws4.Range("W" & FoundCell.Row) And Email_Id.Value = ws4.Range("X" & FoundCell.Row) Then
            'do nothing & exit
        Else
            If Mobile_UAE.Value <> ws4.Range("W" & FoundCell.Row) Or Email_Id.Value <> ws4.Range("X" & FoundCell.Row) Then
                ans = MsgBox("Change in UAE_Mobile_No & Email_ID found. Please confirm to update.", vbYesNo + vbQuestion, "Air Ticket Form")
                If ans = vbYes Then
                    Call Update_record
                End If
            End If
'        Else
            'MsgBox "No changes were made in UAE_Mobile_No & Email_ID Details.", vbExclamation, "Air Ticket Form"
        End If
        
        If ERP_in_Master Is Nothing Then
            'do nothing since ERP not there in master sheet
        Else
            Call TA_Overwrite_Employee_record
            
        End If
        
    End If
End Sub


Sub TA_travel_details_overwrite_check()
    Dim ans As Integer, rgFound As Range, FoundCell As Range, ws7 As Worksheet, ws1 As Worksheet
    
    Set ws7 = Sheet7
    Set ws1 = Sheet1
    Set ERP_No = Sheet1.Range("C5")
    
    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set op1_ref = Sheet1.Range("C19")
    Set op1_1 = Sheet1.Range("B20")
    Set op1_Fare = Sheet1.Range("C21")
    Set op2_ref = Sheet1.Range("F19")
    Set op2_1 = Sheet1.Range("E20")
    Set op2_Fare = Sheet1.Range("F21")
    Set EWF_ref = Sheet1.Range("G4")
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    
    Set FoundCell = ws7.Range("A:A").Find(what:=ERP_No.Value)
    Set rgFound = Sheet4.Range("V:V").Find(Sheet1.Range("C5").Value)
     
    
    If First_Name.Value = ws7.Range("B" & FoundCell.Row) And Date_of_Travel.Value = ws7.Range("C" & FoundCell.Row) And Date_of_Return.Value = ws7.Range("D" & FoundCell.Row) And Boarding_destination.Value = ws7.Range("E" & FoundCell.Row) _
        And Boarding_Port.Value = ws7.Range("F" & FoundCell.Row) And DOT_Time.Value = ws7.Range("G" & FoundCell.Row) _
        And DOR_Time.Value = ws7.Range("H" & FoundCell.Row) And Mobile_Home.Value = ws7.Range("I" & FoundCell.Row) _
        And Mobile_UAE.Value = ws7.Range("J" & FoundCell.Row) And Email_Id.Value = ws7.Range("K" & FoundCell.Row) _
        And op1_ref.Value = ws7.Range("M" & FoundCell.Row) And op2_ref.Value = ws7.Range("N" & FoundCell.Row) _
        And op1_1.Value = ws7.Range("O" & FoundCell.Row) And op2_1.Value = ws7.Range("P" & FoundCell.Row) _
        And op1_Fare.Value = ws7.Range("Q" & FoundCell.Row) And op2_Fare.Value = ws7.Range("R" & FoundCell.Row) And Last_Name.Value = ws7.Range("S" & FoundCell.Row) Then
        'do nothing
    ElseIf First_Name.Value <> ws7.Range("B" & FoundCell.Row) Or Date_of_Travel.Value <> ws7.Range("C" & FoundCell.Row) Or Date_of_Return.Value <> ws7.Range("D" & FoundCell.Row) Or Boarding_destination.Value <> ws7.Range("E" & FoundCell.Row) _
        Or Boarding_Port.Value <> ws7.Range("F" & FoundCell.Row) Or DOT_Time.Value <> ws7.Range("G" & FoundCell.Row) _
        Or DOR_Time.Value <> ws7.Range("H" & FoundCell.Row) Or Mobile_Home.Value <> ws7.Range("I" & FoundCell.Row) _
        Or Mobile_UAE.Value <> ws7.Range("J" & FoundCell.Row) Or Email_Id.Value <> ws7.Range("K" & FoundCell.Row) _
        Or op1_ref.Value <> ws7.Range("M" & FoundCell.Row) Or op2_ref.Value <> ws7.Range("N" & FoundCell.Row) _
        Or op1_1.Value <> ws7.Range("O" & FoundCell.Row) Or op2_1.Value <> ws7.Range("P" & FoundCell.Row) _
        Or op1_Fare.Value <> ws7.Range("Q" & FoundCell.Row) Or op2_Fare.Value <> ws7.Range("R" & FoundCell.Row) Or op2_Fare.Value <> ws7.Range("S" & FoundCell.Row) Then
        
        ans = MsgBox("Change in Travel Details found. Please confirm to update.", vbYesNo + vbQuestion, "Air Ticket Form")
        If ans = vbYes Then
            Call TA_Overwrite_Travel_record
                If rgFound Is Nothing Then 'if erp_no not found in sheet4.col V then exit sub
                    Exit Sub
                Else
                    Call Update_record ' to capture mobile no & email id for future ref
                End If
        Else
            Exit Sub
        End If
    Else
        MsgBox "Employee booking for " & First_Name & " is captured already.", vbExclamation, "Air Ticket Form"
    End If
    
End Sub
Sub TA_Passport_verification()
'If you want to use the Intellisense help showing you the properties
'and methods of the objects as you type you can use Early binding.
'Add a reference to "Microsoft Scripting Runtime" in the VBA editor
'(Tools>References)if you want that.

    Dim FSO As Scripting.FileSystemObject
    Dim FilePath_pdf As String, FilePath_tif As String
    Dim myshape As Shape: Set myshape = Sheet1.Shapes("Button 13")
    
    With myshape
       .ControlFormat.Enabled = False    '---> Disable the button
       .TextFrame.Characters.Font.ColorIndex = 15    '---> Grey out button label 15 to grey out, 1 to enable
    End With
    
    Set FSO = New Scripting.FileSystemObject
    

    FilePath_pdf = "C:\Users\Sidhu\Desktop\ATF\PDF_save\" & Sheet1.Range("C5") & ".pdf"
    FilePath_tif = "C:\Users\Sidhu\Desktop\ATF\PDF_save\" & Sheet1.Range("C5") & ".tif"

    If FSO.FileExists(FilePath_pdf) = False And FSO.FileExists(FilePath_tif) = False Then
        MsgBox "Passport copy is not available. Please contact HR", vbExclamation, "Air Ticket Form"
    Else
        If Dir(FilePath_pdf) <> "" Then
            ActiveWorkbook.FollowHyperlink "C:\Users\Sidhu\Desktop\ATF\PDF_save\" & Sheet1.Range("C5") & ".pdf"
        Else
            ActiveWorkbook.FollowHyperlink "C:\Users\Sidhu\Desktop\ATF\PDF_save\" & Sheet1.Range("C5") & ".tif"
        
        End If
    End If
  
  End Sub

Sub TA_to_HR()
    Dim lastrow As Long, ws As Worksheet, rgFound As Range, RowNo As Integer, ws4 As Worksheet
    Dim check_ERP As Boolean, Row As Long, First_Name As Range
    Dim op1_ref As Range, op1_1 As Range, op1_2 As Range, op1_3 As Range, op1_4 As Range, op1_5 As Range, op1_6 As Range, op1_7 As Range, op1_8 As Range, op1_Fare As Range
    Dim op2_ref As Range, op2_1 As Range, op2_2 As Range, op2_3 As Range, op2_4 As Range, op2_5 As Range, op2_6 As Range, op2_7 As Range, op2_8 As Range, op2_Fare As Range
    
    Set ws = Sheets("To_HR")
    Set ws4 = Sheet4
    Set ERP_No = Sheet1.Range("C5")
    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set op1_ref = Sheet1.Range("C19")
    Set op1_1 = Sheet1.Range("B20")
    Set op1_Fare = Sheet1.Range("C21")
    Set op2_ref = Sheet1.Range("F19")
    Set op2_1 = Sheet1.Range("E20")
    Set op2_Fare = Sheet1.Range("F21")
    Set EWF_ref = Sheet1.Range("G4")
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    
    Row = Row + 1
    
    lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    check_ERP = Not Sheet7.Range("A:A").Find(ERP_No.Value) Is Nothing
    Set rgFound = Sheet7.Range("A:A").Find(ERP_No.Value) 'check ERP no in To_HR sheet
    RowNo = Application.Match(ERP_No.Value, ws4.Range("N:N"), 0)
    
    If IsEmpty(Date_of_Travel) Or IsEmpty(Date_of_Return) Or IsEmpty(Boarding_Port) Or IsEmpty(DOT_Time) _
        Or IsEmpty(DOR_Time) Or IsEmpty(Boarding_destination) Or IsEmpty(Mobile_Home) Or IsEmpty(Mobile_UAE) Or _
         IsEmpty(Email_Id) Or IsEmpty(op1_ref) Or IsEmpty(op2_ref) Or IsEmpty(op2_Fare) Or IsEmpty(op1_1) Or IsEmpty(op2_1) _
         Or IsEmpty(op1_Fare) Or IsEmpty(op2_Fare) Then
        
        MsgBox "Please fill all Travel details.", vbInformation, "Air Ticket Form"
    
    Else
        If rgFound Is Nothing Then 'if erp_no not found in sheet7.col A then do nothing
                'do nothing
        Else
            Call TA_travel_details_overwrite_check ' to overwrite travel record
        End If
        
        If check_ERP Then
                ws4.Range("O" & RowNo) = First_Name.Value
                ws4.Range("P" & RowNo) = Date_of_Travel.Value
                ws4.Range("Q" & RowNo) = Date_of_Return.Value
                ws4.Range("R" & RowNo) = Boarding_destination.Value
            Else
            
                ws.Range("A" & lastrow).Value = ERP_No
                ws.Range("B" & lastrow).Value = First_Name
                ws.Range("C" & lastrow).Value = Format(Date_of_Travel, "dd-mmm-yy")
                ws.Range("D" & lastrow).Value = Format(Date_of_Return, "dd-mmm-yy")
                ws.Range("E" & lastrow).Value = Boarding_destination
                ws.Range("F" & lastrow).Value = Boarding_Port
                ws.Range("G" & lastrow).Value = DOT_Time
                ws.Range("H" & lastrow).Value = DOR_Time
                ws.Range("I" & lastrow).Value = Mobile_Home
                ws.Range("J" & lastrow).Value = Mobile_UAE
                ws.Range("K" & lastrow).Value = Email_Id
                ws.Range("L" & lastrow).Value = EWF_ref
                ws.Range("M" & lastrow).Value = op1_ref
                ws.Range("N" & lastrow).Value = op2_ref
                ws.Range("O" & lastrow).Value = op1_1
                ws.Range("P" & lastrow).Value = op2_1
                ws.Range("Q" & lastrow).Value = op1_Fare
                ws.Range("R" & lastrow).Value = op2_Fare
                ws.Range("S" & lastrow).Value = Last_Name
                
                Call TA_bodyofthe_Mail
                
                MsgBox "Travel detail is saved for HR confirmation.", vbInformation, "Air Ticket Form"
            End If
    
    End If

End Sub

Sub TA_bodyofthe_Mail()
'to capture table to update body of the mail
    Dim lastrow As Long, ws As Worksheet, RowNo As Integer, ws4 As Worksheet
    Dim check_ERP As Boolean, Row As Long, First_Name As Range
    Dim op1_ref As Range, op1_1 As Range, op1_2 As Range, op1_3 As Range, op1_4 As Range, op1_5 As Range, op1_6 As Range, op1_7 As Range, op1_8 As Range, op1_Fare As Range
    Dim op2_ref As Range, op2_1 As Range, op2_2 As Range, op2_3 As Range, op2_4 As Range, op2_5 As Range, op2_6 As Range, op2_7 As Range, op2_8 As Range, op2_Fare As Range
    
    Set ws = Sheets("Consolidate_mail")
    Set ERP_No = Sheet1.Range("C5")
    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set op1_ref = Sheet1.Range("C19")
    Set op1_1 = Sheet1.Range("B20")
    Set op1_Fare = Sheet1.Range("C21")
    Set op2_ref = Sheet1.Range("F19")
    Set op2_1 = Sheet1.Range("E20")
    Set op2_Fare = Sheet1.Range("F21")
    Set EWF_ref = Sheet1.Range("G4")
    Set First_Name = Sheet1.Range("E5")
    
    Row = Row + 1
    
    lastrow = ws.Range("N" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    check_ERP = Not Sheet4.Range("N:N").Find(ERP_No.Value) Is Nothing
    'RowNo = Application.Match(ERP_No.Value, ws4.Range("N:N"), 0)
    'If IsEmpty(Date_of_Travel) Or IsEmpty(Date_of_Return) Or IsEmpty(Boarding_Port) Or IsEmpty(DOT_Time) _
    '    Or IsEmpty(DOR_Time) Or IsEmpty(Boarding_destination) Or IsEmpty(Mobile_Home) Or IsEmpty(Mobile_UAE) Or _
    '     IsEmpty(Email_Id) Or IsEmpty(op1_ref) Or IsEmpty(op2_ref) Or IsEmpty(op2_Fare) Or IsEmpty(op1_1) Or IsEmpty(op2_1) _
    '     Or IsEmpty(op1_Fare) Or IsEmpty(op2_Fare) Then
        
     '   MsgBox "Please fill all Travel details.", vbInformation, "Air Ticket Form"
    
   ' Else
        
            If check_ERP Then
                'ws4.Range("O" & RowNo) = First_Name.Value
                'ws4.Range("P" & RowNo) = Date_of_Travel.Value
                'ws4.Range("Q" & RowNo) = Date_of_Return.Value
                'ws4.Range("R" & RowNo) = Boarding_destination.Value
     
            Else
            
                ws.Range("N" & lastrow).Value = ERP_No
                ws.Range("O" & lastrow).Value = First_Name
                ws.Range("P" & lastrow).Value = Format(Date_of_Travel, "dd-mmm-yy")
                ws.Range("Q" & lastrow).Value = Format(Date_of_Return, "dd-mmm-yy")
                ws.Range("R" & lastrow).Value = Boarding_destination
                'ws.Range("F" & lastrow).Value = Boarding_Port
                'ws.Range("G" & lastrow).Value = DOT_Time
                'ws.Range("H" & lastrow).Value = DOR_Time
                'ws.Range("I" & lastrow).Value = Mobile_Home
                'ws.Range("J" & lastrow).Value = Mobile_UAE
                'ws.Range("K" & lastrow).Value = Email_Id
                'ws.Range("L" & lastrow).Value = EWF_ref
                'ws.Range("M" & lastrow).Value = op1_ref
                'ws.Range("N" & lastrow).Value = op2_ref
                'ws.Range("O" & lastrow).Value = op1_1
                'ws.Range("P" & lastrow).Value = op2_1
                'ws.Range("Q" & lastrow).Value = op1_Fare
                'ws.Range("R" & lastrow).Value = op2_Fare
                
             '   MsgBox "Travel detail is saved for HR confirmation.", vbInformation, "Air Ticket Form"
            End If
    'End If

End Sub
Sub TA_mail_sheet_to_HR()
    Dim lastrow As Long, ws As Worksheet, ws1 As Worksheet, lastrow1 As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim iRet As Integer
    Dim strPrompt As String
    Dim strTitle As String
   
   Application.ScreenUpdating = False
   ActiveWorkbook.Save
   
    Set ws = Sheets("To_HR")
    Set ws1 = Sheets("Consolidate_mail")
    lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
    lastrow1 = ws.Range("N" & Rows.Count).End(xlUp).Row + 1 'Finds the last blank row
   
   Set Sourcewb = ActiveWorkbook
        
If IsEmpty(Sheet7.Range("A2")) Or IsEmpty(Sheet7.Range("B2")) Or IsEmpty(Sheet7.Range("C2")) Or IsEmpty(Sheet7.Range("D2")) Or IsEmpty(Sheet7.Range("E2")) _
        Or IsEmpty(Sheet7.Range("F2")) Or IsEmpty(Sheet7.Range("G2")) Or IsEmpty(Sheet7.Range("H2")) Or IsEmpty(Sheet7.Range("I2")) Or IsEmpty(Sheet7.Range("J2")) _
         Or IsEmpty(Sheet7.Range("K2")) Or IsEmpty(Sheet7.Range("M2")) Or IsEmpty(Sheet7.Range("N2")) Or IsEmpty(Sheet7.Range("O2")) _
         Or IsEmpty(Sheet7.Range("P2")) Or IsEmpty(Sheet7.Range("Q2")) Or IsEmpty(Sheet7.Range("R2")) Then
        
        MsgBox "No data available to send to HR.", vbInformation, "Air Ticket Form"
Else
        
        
        'prompting message whether to send a mail
        strPrompt = "Do you want to send mail to HR?"
        strTitle = "Air Ticket Form"
        iRet = MsgBox(strPrompt, vbYesNo, strTitle)
   
        If iRet = vbNo Then
            Exit Sub
        Else
            'Copy the ActiveSheet to a new workbook
            Sheet7.Copy
            Set Destwb = ActiveWorkbook
        
            'Determine the Excel version and file extension/format
            With Destwb
                If Val(Application.Version) < 12 Then
                    'You use Excel 97-2003
                    FileExtStr = ".xls": FileFormatNum = -4143
                Else
                    'You use Excel 2007-2016
                    Select Case Sourcewb.FileFormat
                    Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
                    Case 52:
                        If .HasVBProject Then
                            FileExtStr = ".xlsm": FileFormatNum = 52
                        Else
                            FileExtStr = ".xlsx": FileFormatNum = 51
                        End If
                    Case 56: FileExtStr = ".xls": FileFormatNum = 56
                    Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
                    End Select
                End If
            End With
        
        
            'Save the new workbook/Mail it/Delete it
            TempFilePath = Environ$("temp") & "\"
            TempFileName = "HR Approval for " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
           
           
           ' Select the range of cells to show in the body of the mail.
            Sheet4.Activate
                 Range("N1:R1" & lastrow1).Select
           
           Sheets("To_HR").Visible = True
                 
                 With Destwb
                 .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
                 On Error Resume Next
                 
                     With ws.MailEnvelope
                         .Introduction = "Please confirm the annual air ticket bookings as per the attached sheet and provide us the LPO No to issue the ticket." 'enter body of the mail
                         .Item.to = "than.gold@gmail.com" ' enter hr email id
                         .Item.Subject = "Annual Air Ticket Booking Confirmations " ' enter subjet of the mail
                         .Item.Attachments.Add Destwb.FullName
                         .Item.send
                     End With
                 End With
            
                 'Application.Wait (Now + TimeValue("0:00:10"))
                 
                 MsgBox "Please confirm OK to forward the Booking Confirmations to HR.", vbInformation, "Air Ticket Form"
                 
                 Application.ScreenUpdating = False
                 
                 'Delete the file you have send
                 Destwb.Close SaveChanges:=False
                 Kill TempFilePath & TempFileName & FileExtStr
                  
                 Sheet7.Activate 'To_HR
                      Range("A2:S2" & lastrow).Delete Shift:=xlUp
                      Range("A1").Select
                Sheet4.Activate 'Consolidate_mail
                      Range("N2:R2" & lastrow).Delete Shift:=xlUp
                      
                Sheet1.Activate
             End If
    End If
End Sub

Sub For_HR()
    MsgBox "For HR purpose only.", vbInformation, "Air Ticket Form"
End Sub

Sub TA_Overwrite_Travel_record()
    Dim RowNo As Integer
    Dim ans As Integer
    Dim ws7 As Worksheet
    
    Set ws7 = Sheet7
    Set ERP_No = Sheet1.Range("C5")
    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set op1_ref = Sheet1.Range("C19")
    Set op1_1 = Sheet1.Range("B20")
    Set op1_Fare = Sheet1.Range("C21")
    Set op2_ref = Sheet1.Range("F19")
    Set op2_1 = Sheet1.Range("E20")
    Set op2_Fare = Sheet1.Range("F21")
    Set EWF_ref = Sheet1.Range("G4")
    Set Last_Name = Sheet1.Range("G5")
    Set First_Name = Sheet1.Range("E5")
    
    RowNo = Application.Match(ERP_No.Value, Sheet7.Range("A:A"), 0)
    
        If ERP_No.Value = Application.VLookup(ERP_No.Value, ws7.Range("A:A"), 1, False) Then
            ws7.Range("B" & RowNo) = First_Name.Value
            ws7.Range("C" & RowNo) = Date_of_Travel.Value
            ws7.Range("D" & RowNo) = Date_of_Return.Value
            ws7.Range("E" & RowNo) = Boarding_destination.Value
            ws7.Range("F" & RowNo) = Boarding_Port.Value
            ws7.Range("G" & RowNo) = DOT_Time.Value
            ws7.Range("H" & RowNo) = DOR_Time.Value
            ws7.Range("I" & RowNo) = Mobile_Home.Value
            ws7.Range("J" & RowNo) = Mobile_UAE.Value
            ws7.Range("K" & RowNo) = Email_Id.Value
            'ws7.Range("L" & RowNo) = EWF_ref.Value
            ws7.Range("M" & RowNo) = op1_ref.Value
            ws7.Range("N" & RowNo) = op2_ref.Value
            ws7.Range("O" & RowNo) = op1_1.Value
            ws7.Range("P" & RowNo) = op2_1.Value
            ws7.Range("Q" & RowNo) = op1_Fare.Value
            ws7.Range("R" & RowNo) = op2_Fare.Value
            ws7.Range("S" & RowNo) = Last_Name.Value
            
            MsgBox "Travel Details record updated successfully.", , "Air Ticket Form"
            
        Else
            'do nothing
        End If
End Sub

Sub ERP_Emp_Lookup()
    
    Dim lookUp_rng As Range, Status As Range
    
    Set ERP_No = Sheet1.Range("C5")
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    Set Job_Desc = Sheet1.Range("C6")
    Set Pay_Grade = Sheet1.Range("E6")
    Set Pay_Start_Date = Sheet1.Range("G6")
    Set Nationality = Sheet1.Range("C7")
    Set Comp_Name = Sheet1.Range("E7")
    Set Dept_Name = Sheet1.Range("G7")
    Set DOB = Sheet1.Range("C8")
    Set BU_No = Sheet1.Range("E8")
    Set BU_Name = Sheet1.Range("G8")
    Set PP_No = Sheet1.Range("C9")
    Set PP_Exp_Date = Sheet1.Range("E9")
    Set Ticket_Destination = Sheet1.Range("G9")
    Set Status = Sheet1.Range("C4")
    
    Set lookUp_rng = Sheet2.Range("B:Y")
    
    If IsError(Application.Match(ERP_No.Value, Sheet2.Range("B:B"), 0)) Then
        'do nothing
    Else
        First_Name.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 23, False)
        Last_Name.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 24, False)
        Job_Desc.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 4, False)
        Pay_Grade.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 5, False)
        Pay_Start_Date.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 7, False)
        Nationality.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 10, False)
        Comp_Name.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 15, False)
        Dept_Name.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 11, False)
        DOB.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 6, False)
        BU_No.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 13, False)
        BU_Name.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 14, False)
        PP_No.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 8, False)
        PP_Exp_Date.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 9, False)
        Ticket_Destination.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 22, False)
        Status.Value = Application.VLookup(ERP_No.Value, lookUp_rng, 21, False)
    End If
End Sub

Sub TA_Overwrite_Employee_record()
    Dim RowNo As Integer
    Dim ans As Integer, FoundCell As Range
    Dim ws2 As Worksheet
    
    Set ws2 = Sheet2
    Set ERP_No = Sheet1.Range("C5")
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    Set Job_Desc = Sheet1.Range("C6")
    Set Pay_Grade = Sheet1.Range("E6")
    Set Pay_Start_Date = Sheet1.Range("G6")
    Set Nationality = Sheet1.Range("C7")
    Set Comp_Name = Sheet1.Range("E7")
    Set Dept_Name = Sheet1.Range("G7")
    Set DOB = Sheet1.Range("C8")
    Set BU_No = Sheet1.Range("E8")
    Set BU_Name = Sheet1.Range("G8")
    Set PP_No = Sheet1.Range("C9")
    Set PP_Exp_Date = Sheet1.Range("E9")
    Set Ticket_Destination = Sheet1.Range("G9")
    
    RowNo = Application.Match(ERP_No.Value, ws2.Range("B:B"), 0)
    Set FoundCell = ws2.Range("B:B").Find(what:=ERP_No.Value)
    
    If First_Name.Value <> ws2.Range("X" & FoundCell.Row) Or Last_Name.Value <> ws2.Range("Y" & FoundCell.Row) Or Job_Desc.Value <> ws2.Range("E" & FoundCell.Row) _
        Or Pay_Grade.Value <> ws2.Range("F" & FoundCell.Row) Or Pay_Start_Date.Value <> ws2.Range("H" & FoundCell.Row) _
        Or Nationality.Value <> ws2.Range("K" & FoundCell.Row) Or Comp_Name.Value <> ws2.Range("P" & FoundCell.Row) _
        Or Dept_Name.Value <> ws2.Range("L" & FoundCell.Row) Or DOB.Value <> ws2.Range("G" & FoundCell.Row) _
        Or BU_No.Value <> ws2.Range("N" & FoundCell.Row) Or BU_Name.Value <> ws2.Range("O" & FoundCell.Row) _
        Or PP_No.Value <> ws2.Range("I" & FoundCell.Row) Or PP_Exp_Date.Value <> ws2.Range("J" & FoundCell.Row) _
        Or Ticket_Destination.Value <> ws2.Range("W" & FoundCell.Row) Then
        
        ans = MsgBox("Change in Employee Details found. Please confirm to update.", vbYesNo + vbQuestion, "Air Ticket Form")
        
        If ans = vbYes Then
            If ERP_No.Value = Application.VLookup(ERP_No.Value, ws2.Range("B:B"), 1, False) Then
                    ws2.Range("X" & RowNo) = First_Name.Value
                    ws2.Range("Y" & RowNo) = Last_Name.Value
                    ws2.Range("E" & RowNo) = Job_Desc.Value
                    ws2.Range("F" & RowNo) = Pay_Grade.Value
                    ws2.Range("H" & RowNo) = Format(Pay_Start_Date.Value, "dd-mmm-yy")
                    ws2.Range("K" & RowNo) = Nationality.Value
                    ws2.Range("P" & RowNo) = Comp_Name.Value
                    ws2.Range("L" & RowNo) = Dept_Name.Value
                    ws2.Range("G" & RowNo) = DOB.Value
                    ws2.Range("N" & RowNo) = BU_No.Value
                    ws2.Range("O" & RowNo) = BU_Name.Value
                    ws2.Range("I" & RowNo) = PP_No.Value
                    ws2.Range("J" & RowNo) = Format(PP_Exp_Date.Value, "dd-mmm-yy")
                    ws2.Range("W" & RowNo) = Ticket_Destination.Value
                    
                    MsgBox "Employee Details record updated successfully.", , "Air Ticket Form"
                    
                Else
                    'ERP no not found in master file
                End If
        Else
            ' ans is no then don't update
        End If
    Else
        MsgBox "There is no change Employee details of " & First_Name, vbExclamation, "Air Ticket Form"
    End If
    
End Sub
Sub check()
Dim FoundCell As Range
Dim ws7 As Worksheet

    Set Date_of_Travel = Sheet1.Range("C12")
    Set Date_of_Return = Sheet1.Range("E12")
    Set Boarding_Port = Sheet1.Range("G12")
    Set DOT_Time = Sheet1.Range("C13")
    Set DOR_Time = Sheet1.Range("E13")
    Set Boarding_destination = Sheet1.Range("G13")
    Set Mobile_Home = Sheet1.Range("C14")
    Set Mobile_UAE = Sheet1.Range("E14")
    Set Email_Id = Sheet1.Range("G14")
    Set op1_ref = Sheet1.Range("C19")
    Set op1_1 = Sheet1.Range("B20")
    Set op1_Fare = Sheet1.Range("C21")
    Set op2_ref = Sheet1.Range("F19")
    Set op2_1 = Sheet1.Range("E20")
    Set op2_Fare = Sheet1.Range("F21")
    Set EWF_ref = Sheet1.Range("G4")
    Set First_Name = Sheet1.Range("E5")
    Set Last_Name = Sheet1.Range("G5")
    Set ERP_No = Sheet1.Range("C5")
Set ws7 = Sheet7
Set FoundCell = ws7.Range("A:A").Find(what:=ERP_No.Value)

If Date_of_Travel.Value = ws7.Range("C" & FoundCell.Row) And Date_of_Return.Value = ws7.Range("D" & FoundCell.Row) And Boarding_destination.Value = ws7.Range("E" & FoundCell.Row) _
        And Boarding_Port.Value = ws7.Range("F" & FoundCell.Row) And DOT_Time.Value = ws7.Range("G" & FoundCell.Row) _
        And DOR_Time.Value = ws7.Range("H" & FoundCell.Row) And Mobile_Home.Value = ws7.Range("I" & FoundCell.Row) _
        And Mobile_UAE.Value = ws7.Range("J" & FoundCell.Row) And Email_Id.Value = ws7.Range("K" & FoundCell.Row) _
        And op1_ref.Value = ws7.Range("M" & FoundCell.Row) And op2_ref.Value = ws7.Range("N" & FoundCell.Row) _
        And op1_1.Value = ws7.Range("O" & FoundCell.Row) And op2_1.Value = ws7.Range("P" & FoundCell.Row) _
        And op1_Fare.Value = ws7.Range("Q" & FoundCell.Row) And op2_Fare.Value = ws7.Range("R" & FoundCell.Row) And Last_Name.Value = ws7.Range("S" & FoundCell.Row) Then
    
    MsgBox "No Change in data", vbExclamation, "Air Ticket Form"

Else
    MsgBox "Change in data", vbExclamation, "Air Ticket Form"
End If
End Sub
