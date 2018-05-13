VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GWDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_objDatabase As DAO.Database

    
' ----------------------------------------------------------------
' Procedure Name: Class_Initialize
' Purpose: Calls CheckDatabase Sub to verify DB integrity
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub Class_Initialize()

    ' * ' Check For Globals
    If g_objGWClient Is Nothing Then
        Set g_objGWClient = New GymWizardClient
    End If
    
    ' * ' Check that Database is Ready
    Call CheckDatabase
    
    ' * ' Check For Log In
    If Not SetUpCheck Then
        Load frmLogInSetUp
        frmLogInSetUp.Show
        
    End If
    
    
End Sub
Public Function SetUpCheck() As Boolean
    
    Dim objRecordset As DAO.Recordset
    Set objRecordset = m_objDatabase.OpenRecordset("SELECT [Value] FROM [User Info] WHERE [Variable] = 'Admin_UserName';")
    
    If (objRecordset.EOF And objRecordset.BOF) Or objRecordset![Value] = " " Then
        SetUpCheck = False
    Else:
        SetUpCheck = True
    End If
    
End Function
Public Sub FormatClasses()

    Dim objRecordset    As DAO.Recordset
    Dim sInvoiceType    As String
    
    Set objRecordset = Database.OpenRecordset("SELECT * FROM [Scheduled Classes];")
    
    With objRecordset
        .MoveFirst
        
        Do Until .EOF
            Dim s_TitleNew As String
            
            .Edit
            
            s_TitleNew = ![Raw Title]
            
            ' * ' BJJ Classes need to be shortened to the BJJ shorthand as well as determine if they are no-gi
            If InStr(![Raw Title], "Jiu-Jitsu") Then
                
                ' * ' Check IF Women's
                If InStr(![Raw Title], "Women's") Then
                    s_TitleNew = "Women's BJJ"
                Else:
                    s_TitleNew = "BJJ"
                End If
                
                ' * ' Check If No GI
                If ![Special] = "No-GI" Or InStr(![Raw Title], "No-Gi") Then
                    s_TitleNew = s_TitleNew & " (No-GI)"
                End If
                
                s_TitleNew = "[C] " & s_TitleNew
                
            End If
            
            ' * ' Womens BJJ
            If InStr(![Raw Title], "Women's BJJ") Or InStr(![Raw Title], "Women's Brazilian") Then
                
                s_TitleNew = "[C] Women's BJJ"
                
                If InStr(![Raw Title], "No-Gi") Then
                    s_TitleNew = s_TitleNew & " (No-GI)"
                End If
            End If
            ' * ' YOUTH JIU JITSU
            If InStr(![Raw Title], "Youth Brazilian") Then
                
                ' * ' Comp?
                If InStr(![Raw Title], "Advanced") Then
                    s_TitleNew = "[C] Youth BJJ - Advanced"
                Else:
                    s_TitleNew = "Youth BJJ"
                End If
                
            End If
            
            
            ' * ' SPECIAL CLASES
            ' * ' Some classes are special event classes and are formatted with a SPECIAL 60 MIN Event Title and a subtitle for the actual class
            If InStr(![Raw Title], "SPECIAL") Then
                a_split = Split(![Raw Title], "<BR>")
                s_TitleNew = "*SPECIAL* " & a_split(1)
            End If
            
            
            ' * ' Youth DUT
            If InStr(![Raw Title], "SAQ") Then
                s_TitleNew = "S.A.Q"
            End If
            
            ' * ' Little Graps
            If InStr(![Raw Title], "Lil' Grapplers") Then
                s_TitleNew = "Lil' Grapplers"
            End If
            
            ' * ' Family DUT
            If InStr(![Raw Title], "Family Daily") Then
                s_TitleNew = "[B] Family DUT"
            End If
            
            ' * ' High Octane gets Cut OFF
            If InStr(![Raw Title], "High Octane") Then
                s_TitleNew = "High Octane Conditioning"
            End If
            
            ' * ' Box SKill and Tech
            If InStr(![Raw Title], "Boxing Skills") Then
                s_TitleNew = "Boxing Skills & Technique"
            End If
            
            ' * ' DUT
            If InStr(![Raw Title], "Daily Ultimate Train") Then
                s_TitleNew = "DUT"
            End If
            
            
            ' * ' Room Designations
            Select Case s_TitleNew
                
                Case "Striking", "Fight Fit", "BAGWORX", "Boxing Skills & Technique", "TRX", "MATRX":
                    s_TitleNew = "[A] " & s_TitleNew
                
                Case "DUT":
                    s_TitleNew = "[B] " & s_TitleNew
                
                Case "Judo", "Muay Thai Level 1", "MMA":
                    s_TitleNew = "[O] " & s_TitleNew
            End Select
                
            ' * ' Add To Record
            ![Formatted Title] = s_TitleNew
            .Update
            
            ' * ' Loop Next
            .MoveNext
        Loop
        
    End With
    
                
            
End Sub
' ----------------------------------------------------------------
' Procedure Name: ProcessRefundData
' Purpose: Analyzes data in Refund table to indentify Invoice Types
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Public Sub ProcessRefundData()
    Dim objRecordset    As DAO.Recordset
    Dim sInvoiceType    As String
    
    Set objRecordset = Database.OpenRecordset("SELECT * FROM [Refunds] WHERE [Invoice Type] is Null Or [Invoice Type] = 'Unknown';")
    
    With objRecordset
        .MoveFirst
        
        ' * ' Loop Through All Records
        Do Until .EOF
            
            If IsNull(![Purchase Description]) Then
                sInvoiceType = "Product"
            ElseIf ![Original Payment] = "Down" Then
                sInvoiceType = "Down Payment"
            ElseIf InStr(1, ![Purchase Description], "PT/MMA") > 0 Then
                sInvoiceType = "PT"
            ElseIf InStr(1, ![Purchase Description], "Per Week") > 0 Then
                sInvoiceType = "PT"
            ElseIf InStr(1, ![Purchase Description], "Dues") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Ultimate") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Fitness") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Youth") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Champion") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Late Fee") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Signature Enhancement") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Freeze") > 0 Then
                sInvoiceType = "Dues"
            ElseIf InStr(1, ![Purchase Description], "Returned Payment Fee") > 0 Then
                sInvoiceType = "Dues"
            Else:
                sInvoiceType = "Unknown"
            End If
            
            ' * ' Edit Record
            .Edit
            ![Invoice Type] = sInvoiceType
            .Update
            .MoveNext
        
        Loop
                
    End With
                
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: Class_Terminate
' Purpose: Cleans Up
' Procedure Kind: Destructor (Terminate)
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub Class_Terminate()
    Database.Close
End Sub

Private Property Get Database() As DAO.Database

    Set Database = m_objDatabase

End Property

' ----------------------------------------------------------------
' Procedure Name: CheckDatabase
' Purpose: Check Database EXISTS and is Updated
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CheckDatabase()
    
    ' * ' Check it Exists
    If Not FileExists(g_objGWClient.DatabasePath) Then
        CreateDatabase
    Else:
        ' * ' OPEN DB
        Set m_objDatabase = OpenDatabase(g_objGWClient.DatabasePath)
        
        ' * ' Check All Tables
        If Not TableExists("Gross Sales Detail") Then
            CreateGSDTable
        End If
        
        If Not TableExists("Refunds") Then
            CreateRefundsTable
        End If
        
        If Not TableExists("Invoices Coming Due") Then
            CreateICDTable
        End If
        
        If Not TableExists("Club Past Due") Then
            CreateCPDTable
        End If
        
        If Not TableExists("Agreements Log") Then
            CreateAgreementTable
        End If
        
        If Not TableExists("Agreement Statement") Then
            CreateAgreementStatementTable
        End If
        
        If Not TableExists("Scheduled Classes") Then
            CreateScheduledClassesTable
        End If
        
        If Not TableExists("Member Export") Then
            CreateMemExpTable
        End If
        
        If Not TableExists("Attendance") Then
            CreateAttTable
        End If
        
        If Not TableExists("Frozen Members") Then
            CreateFrozenMemExpTable
        End If
        
        If Not TableExists("User Info") Then
            CreateUserTable
            
            Dim objRecordset As DAO.Recordset
            m_objDatabase.Execute ("INSERT INTO [User Info] ([Variable],[Value]) VALUES ('Admin_UserName', ' ');")
            m_objDatabase.Execute ("INSERT INTO [User Info] ([Variable],[Value]) VALUES ('Admin_Password', ' ');")
            
        End If
        
        If Not TableExists("Package Details") Then
            CreatePackageTypeTable
        End If
    End If
    
    
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: CreateCPDTable
' Purpose: Creates Table for Club Past Due Report
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CreateCPDTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Club Past Due")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Location", dbText)
        .Fields.Append .CreateField("User ID", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("Email Address", dbText)
        .Fields.Append .CreateField("Home Phone", dbText)
        .Fields.Append .CreateField("Cell Phone", dbText)
        .Fields.Append .CreateField("Work Phone", dbText)
        .Fields.Append .CreateField("Original Agreement Date", dbDate)
        .Fields.Append .CreateField("Detail", dbText)
        .Fields.Append .CreateField("Amount", dbCurrency)
        .Fields.Append .CreateField("Sales Tax", dbCurrency)
        .Fields.Append .CreateField("Due Date", dbDate)
        .Fields.Append .CreateField("Age Days", dbText)
        .Fields.Append .CreateField("In Collections Status?", dbBoolean)
        .Fields.Append .CreateField("Has Payment Details on File?", dbBoolean)
        .Fields.Append .CreateField("Attempts", dbText)
        .Fields.Append .CreateField("Is Return?", dbBoolean)
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
    ' ' Clean it up
    Set objTableDef = Database.TableDefs("Club Past Due")
    With objTableDef
        .Fields("Has Payment Details on File?").Properties.Append .Fields("Has Payment Details on File?").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("In Collections Status?").Properties.Append .Fields("In Collections Status?").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("Is Return?").Properties.Append .Fields("Is Return?").CreateProperty("DisplayControl", dbInteger, 106)
    End With
    
    ' ' set nothing
    Set objTableDef = Nothing
    
End Sub
Private Sub CreatePackageTypeTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Package Details")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Package Name", dbText)
        .Fields.Append .CreateField("Package Type", dbText)
        .Fields.Append .CreateField("Duration", dbText)
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
        
End Sub
Private Sub CreateAgreementStatementTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Agreement Statement")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("User ID", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("InvoiceID", dbText)
        .Fields.Append .CreateField("Amount", dbCurrency)
        .Fields.Append .CreateField("Sales Tax", dbCurrency)
        .Fields.Append .CreateField("Detail", dbText)
        .Fields.Append .CreateField("Due Date", dbDate)
        .Fields.Append .CreateField("Paid Date", dbDate)
        .Fields.Append .CreateField("Payment Method", dbText)
        .Fields.Append .CreateField("Primary Sales Person", dbText)
        .Fields.Append .CreateField("Secondary Sales Person", dbText)
        .Fields.Append .CreateField("Note", dbText)
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub

Public Function GetRecordset(sQuery As String) As DAO.Recordset

    Set GetRecordset = Database.OpenRecordset(sQuery)
End Function
Public Sub ReBuildTable(sTableName As String)
    Database.TableDefs.Delete (sTableName)
    CheckDatabase
End Sub
Private Sub CreateScheduledClassesTable()

    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Scheduled Classes")
    
    With objTableDef
        .Fields.Append .CreateField("WeekDay", dbText)
        .Fields.Append .CreateField("Start Time", dbDate)
        .Fields.Append .CreateField("Raw Title", dbText)
        .Fields.Append .CreateField("Formatted Title", dbText)
        .Fields.Append .CreateField("Coach1", dbText)
        .Fields.Append .CreateField("Coach2", dbText)
        .Fields.Append .CreateField("Class ID", dbText)
        .Fields.Append .CreateField("Special", dbText) '' Used for NO-Gi Variant BJJ Classes
        .Fields.Append .CreateField("Type", dbText)
        
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
Private Sub CreateAgreementTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Agreements Log")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("Current Status", dbText)
        .Fields.Append .CreateField("Membership Expires", dbText)
        .Fields.Append .CreateField("Agreement ID", dbText)
        .Fields.Append .CreateField("Division", dbText)
        .Fields.Append .CreateField("District", dbText)
        .Fields.Append .CreateField("Location", dbText)
        .Fields.Append .CreateField("State", dbText)
        .Fields.Append .CreateField("User ID", dbText)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("Phone", dbText)
        .Fields.Append .CreateField("Cell Phone", dbText)
        .Fields.Append .CreateField("Email", dbText)
        .Fields.Append .CreateField("Payment Preference", dbText)
        .Fields.Append .CreateField("Description", dbText)
        .Fields.Append .CreateField("Cancelled Date", dbDate)
        .Fields.Append .CreateField("Cancelled By", dbText)
        .Fields.Append .CreateField("Total Sessions", dbText)
        .Fields.Append .CreateField("Total Free Sessions", dbText)
        .Fields.Append .CreateField("Sales Person", dbText)
        .Fields.Append .CreateField("Second Sales Person", dbText)
        .Fields.Append .CreateField("First Invoice Due", dbDate)
        .Fields.Append .CreateField("Total Invoice", dbCurrency)
        .Fields.Append .CreateField("Setup Fee", dbCurrency)
        .Fields.Append .CreateField("Paid Upfront", dbCurrency)
        .Fields.Append .CreateField("Enhancement Fee Amount", dbCurrency)
        .Fields.Append .CreateField("Enhancement Fee Date", dbDate)
        .Fields.Append .CreateField("Initial Planned Draft", dbCurrency)
        .Fields.Append .CreateField("Signature", dbText)
        .Fields.Append .CreateField("Total Non Refunded Payments Made To-Date", dbCurrency)
        .Fields.Append .CreateField("Total Non Refunded Payments Made During Report Range", dbCurrency)
        .Fields.Append .CreateField("Total Refunded", dbCurrency)
        .Fields.Append .CreateField("Total Amount Remaining Including Cancelled Excluding Evergreen", dbCurrency)
        .Fields.Append .CreateField("Past Due Remaining EFT Excluding Evergreen", dbCurrency)
        .Fields.Append .CreateField("Cancelled EFT Excluding Evergreen", dbCurrency)
        .Fields.Append .CreateField("Unpaid Evergeen Remaining EFT", dbCurrency)
        .Fields.Append .CreateField("New EFT", dbCurrency)
        .Fields.Append .CreateField("Amenities", dbText)
        .Fields.Append .CreateField("Agreement Review Status", dbText)
        .Fields.Append .CreateField("Agreement Notes", dbText)
        .Fields.Append .CreateField("Number of Previous Agreements", dbText)
        .Fields.Append .CreateField("Previous Agreement Date", dbDate)
        .Fields.Append .CreateField("Previous Agreement Detail", dbText)
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
' ----------------------------------------------------------------
' Procedure Name: CreateGSDTable
' Purpose: Creates Table for Gross Sales Detail Report
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CreateGSDTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Gross Sales Detail")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("UserID", dbText)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("InvoiceID", dbText)
        .Fields.Append .CreateField("Due Date", dbDate)
        .Fields.Append .CreateField("Type", dbText)
        .Fields.Append .CreateField("Payment Type", dbText)
        .Fields.Append .CreateField("Amount", dbCurrency)
        .Fields.Append .CreateField("Sales Tax", dbCurrency)
        .Fields.Append .CreateField("Total", dbCurrency)
        .Fields.Append .CreateField("List Price", dbCurrency)
        .Fields.Append .CreateField("Discount", dbCurrency)
        .Fields.Append .CreateField("Wholesale Price", dbCurrency)
        .Fields.Append .CreateField("Transaction", dbText)
        .Fields.Append .CreateField("Category", dbText)
        .Fields.Append .CreateField("Detail", dbText)
        .Fields.Append .CreateField("Product Category", dbText)
        .Fields.Append .CreateField("Session Credits", dbText)
        .Fields.Append .CreateField("Primary Sales Person", dbText)
        .Fields.Append .CreateField("Secondary Sales Person", dbText)
        .Fields.Append .CreateField("Ran By", dbText)
        .Fields.Append .CreateField("Ran By Name", dbText)
        .Fields.Append .CreateField("POS Terminal", dbText)
        .Fields.Append .CreateField("Note", dbText)
        .Fields.Append .CreateField("Invoice Category", dbText)
        .Fields.Append .CreateField("Invoice Class", dbText)
        .Fields.Append .CreateField("Invoice Type", dbText)
        .Fields.Append .CreateField("GSL Category", dbText)
        .Fields.Append .CreateField("Due Status", dbText)
        .Fields.Append .CreateField("Processed By", dbText)
        .Fields.Append .CreateField("Invoice Source", dbText)
        .Fields.Append .CreateField("Agreement ID", dbText)
        .Fields.Append .CreateField("Agreement Sales Person 1", dbText)
        .Fields.Append .CreateField("Agreement Sales Person 2", dbText)
        .Fields.Append .CreateField("Agreement Date", dbDate)
    End With
    
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
Private Sub CreateMemExpTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Member Export")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Member First Name", dbText)
        .Fields.Append .CreateField("Member Last Name", dbText)
        .Fields.Append .CreateField("Member ID", dbText)
        .Fields.Append .CreateField("Status", dbText)
        .Fields.Append .CreateField("Member Since", dbDate)
        .Fields.Append .CreateField("Membership Expires", dbDate)
        .Fields.Append .CreateField("Membership Ended", dbDate)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Pincode", dbText)
        .Fields.Append .CreateField("Membership Type", dbText)
        .Fields.Append .CreateField("Internal Lead Type", dbText)
        .Fields.Append .CreateField("Referral Type", dbText)
        .Fields.Append .CreateField("Gender", dbText)
        .Fields.Append .CreateField("DOB", dbDate)
        .Fields.Append .CreateField("Address", dbText)
        .Fields.Append .CreateField("City", dbText)
        .Fields.Append .CreateField("State", dbText)
        .Fields.Append .CreateField("Zip", dbText)
        .Fields.Append .CreateField("Email", dbText)
        .Fields.Append .CreateField("Phone", dbText)
        .Fields.Append .CreateField("Cell Phone", dbText)
        .Fields.Append .CreateField("Work Phone", dbText)
        .Fields.Append .CreateField("Emergency Contact", dbText)
        .Fields.Append .CreateField("Emergency Contact Number", dbText)
        .Fields.Append .CreateField("Current Past Due", dbCurrency)
        .Fields.Append .CreateField("Last Payment Date", dbDate)
        .Fields.Append .CreateField("Last Payment Amount", dbCurrency)
        .Fields.Append .CreateField("Due Next 30 Days", dbCurrency)
        .Fields.Append .CreateField("Next Draft Date", dbDate)
        .Fields.Append .CreateField("Card Exp", dbText)
        .Fields.Append .CreateField("Alerts", dbText)
    End With
    
    For Each objField In objTableDef.Fields
        objField.AllowZeroLength = True
    Next
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
Private Sub CreateFrozenMemExpTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Frozen Members")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Member First Name", dbText)
        .Fields.Append .CreateField("Member Last Name", dbText)
        .Fields.Append .CreateField("Member ID", dbText)
        .Fields.Append .CreateField("Status", dbText)
        .Fields.Append .CreateField("Member Since", dbDate)
        .Fields.Append .CreateField("Membership Expires", dbDate)
        .Fields.Append .CreateField("Membership Ended", dbDate)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Pincode", dbText)
        .Fields.Append .CreateField("Membership Type", dbText)
        .Fields.Append .CreateField("Internal Lead Type", dbText)
        .Fields.Append .CreateField("Referral Type", dbText)
        .Fields.Append .CreateField("Gender", dbText)
        .Fields.Append .CreateField("DOB", dbDate)
        .Fields.Append .CreateField("Address", dbText)
        .Fields.Append .CreateField("City", dbText)
        .Fields.Append .CreateField("State", dbText)
        .Fields.Append .CreateField("Zip", dbText)
        .Fields.Append .CreateField("Email", dbText)
        .Fields.Append .CreateField("Phone", dbText)
        .Fields.Append .CreateField("Cell Phone", dbText)
        .Fields.Append .CreateField("Work Phone", dbText)
        .Fields.Append .CreateField("Emergency Contact", dbText)
        .Fields.Append .CreateField("Emergency Contact Number", dbText)
        .Fields.Append .CreateField("Current Past Due", dbCurrency)
        .Fields.Append .CreateField("Last Payment Date", dbDate)
        .Fields.Append .CreateField("Last Payment Amount", dbCurrency)
        .Fields.Append .CreateField("Due Next 30 Days", dbCurrency)
        .Fields.Append .CreateField("Next Draft Date", dbDate)
        .Fields.Append .CreateField("Card Exp", dbText)
        .Fields.Append .CreateField("Alerts", dbText)
    End With
    
    For Each objField In objTableDef.Fields
        objField.AllowZeroLength = True
    Next
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
Private Sub CreateAttTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Attendance")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Club Name", dbText)
        .Fields.Append .CreateField("UserID", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("Check In Date / Time", dbDate)
        .Fields.Append .CreateField("Home Club", dbText)
        .Fields.Append .CreateField("Allowed", dbText)
        .Fields.Append .CreateField("Denied", dbText)
        .Fields.Append .CreateField("Gender", dbText)
        .Fields.Append .CreateField("DOB", dbDate)
        .Fields.Append .CreateField("Membership Type", dbText)
        .Fields.Append .CreateField("Prospect Type", dbText)
        .Fields.Append .CreateField("Responsible Person Last Name", dbText)
        .Fields.Append .CreateField("Responsible Person First Name", dbText)
        .Fields.Append .CreateField("Custom Field", dbText)
    End With
    
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
Private Sub CreateUserTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("User Info")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Variable", dbText)
        .Fields.Append .CreateField("Value", dbText)
    End With
    
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
' ----------------------------------------------------------------
' Procedure Name: CreateRefundsTable
' Purpose: Create Table for Refunds Report
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CreateRefundsTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Dim objField As DAO.Field
    Set objTableDef = Database.CreateTableDef("Refunds")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("Original Payment Date", dbDate)
        .Fields.Append .CreateField("Amount", dbCurrency)
        .Fields.Append .CreateField("Tax", dbCurrency)
        .Fields.Append .CreateField("Total Amount", dbCurrency)
        .Fields.Append .CreateField("Original Payment ID", dbText)
        .Fields.Append .CreateField("Purchase Description", dbText)
        .Fields.Append .CreateField("Reason for Refund", dbText)
        .Fields.Append .CreateField("Original Payment", dbText)
        .Fields.Append .CreateField("Payment Type", dbText)
        .Fields.Append .CreateField("Card Entry Method", dbText)
        .Fields.Append .CreateField("User ID", dbText)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("Detail", dbText)
        .Fields.Append .CreateField("Refunded By", dbText)
        .Fields.Append .CreateField("Original Ran By", dbText)
        .Fields.Append .CreateField("Invoice Type", dbText)
    End With
    
    For Each objField In objTableDef.Fields
        objField.AllowZeroLength = True
    Next
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: CreateTransLedgerTable
' Purpose: Creates Table for Transaction Ledger Report
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CreateTransLedgerTable()
    
    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Transaction Ledger")
    
    ' ' Set Fields
    With objTableDef
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("Transaction Detail", dbText)
        .Fields.Append .CreateField("Transaction", dbCurrency)
        .Fields.Append .CreateField("Customer Balance", dbCurrency)
    End With
    
    ' ' Add Table To DB
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
End Sub
' ----------------------------------------------------------------
' Procedure Name: CreateDatabase
' Purpose: Creates DB File In Gym Wizard Folder and Populates it with Tables
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Public Sub CreateDatabase()

    Dim objCatalog As New ADOX.Catalog
    
    ' * ' Create DB File
    objCatalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & g_objGWClient.DatabasePath
    
    ' * ' Now OPEN DATABASE
    Set m_objDatabase = OpenDatabase(g_objGWClient.DatabasePath)
    
    ' * ' Populate Tables
    CheckDatabase
    
End Sub

    
' ----------------------------------------------------------------
' Procedure Name: CreateICDTable
' Purpose: Creates Table for Invoices Coming Due Report
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub CreateICDTable()

    ' ' Create New Table Def
    Dim objTableDef As DAO.TableDef
    Set objTableDef = Database.CreateTableDef("Invoices Coming Due")
    
    With objTableDef
        .Fields.Append .CreateField("Invoice Number", dbText)
        .Fields.Append .CreateField("Amount", dbCurrency)
        .Fields.Append .CreateField("Sales Tax Amount", dbCurrency)
        .Fields.Append .CreateField("Due Date", dbDate)
        .Fields.Append .CreateField("Paid Date", dbDate)
        .Fields.Append .CreateField("Scheduled Cancel Date", dbDate)
        .Fields.Append .CreateField("Agreement Number", dbText)
        .Fields.Append .CreateField("Last Name", dbText)
        .Fields.Append .CreateField("First Name", dbText)
        .Fields.Append .CreateField("User ID", dbText)
        .Fields.Append .CreateField("User Pay Preference", dbText)
        .Fields.Append .CreateField("Has CC", dbBoolean)
        .Fields.Append .CreateField("Has ACH", dbBoolean)
        .Fields.Append .CreateField("Barcode", dbText)
        .Fields.Append .CreateField("Location", dbText)
        .Fields.Append .CreateField("Details", dbText)
        .Fields.Append .CreateField("Original Agreement Date", dbDate)
        .Fields.Append .CreateField("Sold By", dbText)
        .Fields.Append .CreateField("Home Phone", dbText)
        .Fields.Append .CreateField("Cell Phone", dbText)
        .Fields.Append .CreateField("Work Phone", dbText)
        .Fields.Append .CreateField("Email", dbText)
        .Fields.Append .CreateField("Invoice Category", dbText)
        .Fields.Append .CreateField("Invoice Class", dbText)
        .Fields.Append .CreateField("Invoice Type", dbText)
        .Fields.Append .CreateField("GSL Category", dbText)
        .Fields.Append .CreateField("Due Status", dbText)
        .Fields.Append .CreateField("Processed By", dbText)
        .Fields.Append .CreateField("Invoice Source", dbText)
        .Fields.Append .CreateField("Note", dbInteger)
    End With
    
    Database.TableDefs.Append objTableDef
    Set objTableDef = Nothing
    
    ' * ' Edit CC ACH Bool Fields
    Set objTableDef = Database.TableDefs("Invoices Coming Due")
    With objTableDef
        .Fields("Has CC").Properties.Append .Fields("Has CC").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("Has ACH").Properties.Append .Fields("Has ACH").CreateProperty("DisplayControl", dbInteger, 106)
    End With
    
    Set objTableDef = Nothing
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: TableExists
' Purpose: Checks if Table exists in GWDatabase
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sTableName (String):
' Return Type: Boolean
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Function TableExists(ByVal sTableName As String) As Boolean
    Dim bResult     As Boolean
    Dim objTableDef As DAO.TableDef
    
    
    On Error Resume Next
    Set objTableDef = Database.TableDefs(sTableName)
    
    bResult = (Err.Number = 0)
    Err.Clear
    
    TableExists = bResult
End Function
