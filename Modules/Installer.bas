Attribute VB_Name = "Installer"
Sub Install()

    Dim FSO As New FileSystemObject
        
    ' * ' Saved Settings
    mStr_MainDir = Range("MAINDIR").Value
    mStr_ScheduleDir = Range("SCHEDDIR").Value
    mStr_TapDir = Range("TAPDIR").Value
    
    ' * ' Generate Paths
    mStr_DesktopDir = Environ("USERPROFILE") & "\Desktop"
    mStr_DocumentsDir = Environ("USERPROFILE") & "\My Documents"
    
    
    mStr_MainDir2 = mStr_DocumentsDir & "\Gym Wizard"
    mStr_DownloadDir2 = mStr_DocumentsDir & "\Gym Wizard\Downloads"
    mStr_ScheduleDir2 = mStr_DesktopDir & "\Class Schedules"
    mStr_TapDir2 = mStr_DesktopDir & "\Tap Lists"
    mStr_DBPath = mStr_MainDir2 & "\UFCDB.accdb"

    ' * ' Check For Dir
    If Not FSO.FolderExists(mStr_MainDir2) Then
        FSO.CreateFolder mStr_MainDir2
        FSO.CreateFolder mStr_DownloadDir2
    End If
    
    If Not FSO.FolderExists(mStr_ScheduleDir2) Then
        FSO.CreateFolder mStr_ScheduleDir2
    End If
    
    If Not FSO.FolderExists(mStr_TapDir2) Then
        FSO.CreateFolder mStr_TapDir2
    End If
    
    If Not FSO.FileExists(mStr_DBPath) Then
    
        ' ' Create DB
        Set Catalog = New ADOX.Catalog
        Catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mStr_DBPath
        
        ' * ' Populate
        CreateDatabase mStr_DBPath
        
    End If
    
    ' * ' Set Settings
    Range("MAINDIR").Value = mStr_MainDir2
    Range("SCHEDDIR").Value = mStr_ScheduleDir2
    Range("TAPDIR").Value = mStr_TapDir2
    Range("DBDIR").Value = mStr_DBPath
    
   
    
End Sub

'=======================================================================================
' Method   : AttendanceReportButton_Click
' Author   : Matthew Rodriguez
' Created  : 1/27/2017
' Purpose  :
' Called by: user
' Arguments: Creates Proper Database For Usage Later
' Comments : Selected file must be Daily Attendance
' Changes-------------------------------------------------------------------------------
' Date            Programmer                Change
'
'=======================================================================================
Function CreateDatabase(ByVal Path As String) As Boolean

    Dim Catalog As ADOX.Catalog
    Dim db As DAO.Database
    Dim fl As DAO.Field
    Dim pr As DAO.Property
    Dim td As DAO.TableDef
    Dim tdfNew As DAO.TableDef
    
    On Error GoTo Handler
    
        Set db = OpenDatabase(Path)
        Set tdfNew = db.CreateTableDef("Club Past Due")
        
        With tdfNew
            .Fields.Append .CreateField("Location", dbText)
            .Fields.Append .CreateField("User ID", dbInteger)
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
            .Fields.Append .CreateField("Age Days", dbInteger)
            .Fields.Append .CreateField("In Collections Status?", dbBoolean)
            .Fields.Append .CreateField("Has Payment Details on File?", dbBoolean)
            .Fields.Append .CreateField("Attempts", dbInteger)
            .Fields.Append .CreateField("Is Return?", dbBoolean)
            
            
            
        End With
        
        
        ' ' Add To DB
        db.TableDefs.Append tdfNew
        Set tdfNew = Nothing
        
        Set tdfNew = db.TableDefs("Club Past Due")
        
        With tdfNew
            .Fields("Has Payment Details on File?").Properties.Append .Fields("Has Payment Details on File?").CreateProperty("DisplayControl", dbInteger, 106)
            .Fields("In Collections Status?").Properties.Append .Fields("In Collections Status?").CreateProperty("DisplayControl", dbInteger, 106)
            .Fields("Is Return?").Properties.Append .Fields("Is Return?").CreateProperty("DisplayControl", dbInteger, 106)
        End With
        
        
        Set tdfNew = Nothing
        Set tdfNew = db.CreateTableDef("Invoices Coming Due")
        
        
        With tdfNew
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
        End With
        
        db.TableDefs.Append tdfNew
        Set tdfNew = Nothing
        
        Set tdfNew = db.TableDefs("Invoices Coming Due")
        
        With tdfNew
            .Fields("Has CC").Properties.Append .Fields("Has CC").CreateProperty("DisplayControl", dbInteger, 106)
            .Fields("Has ACH").Properties.Append .Fields("Has ACH").CreateProperty("DisplayControl", dbInteger, 106)
        End With
        
        Set tdfNew = Nothing
        Set tdfNew = db.CreateTableDef("Cancelled Agreements")
        
        ' ' Invoice Number is Primary Key
        Set idx = tdfNew.CreateIndex("PrimaryKey")
        
        With idx
            .Name = "PrimaryKey"
            .Primary = True
            .Required = True
            .IgnoreNulls = False
        End With
        
        Set fld = idx.CreateField("Agreement Id", dbText)
        idx.Fields.Append fld
        
        tdfNew.Indexes.Append idx
        
        With tdfNew
            .Fields.Append .CreateField("Cancel Date", dbDate)
            .Fields.Append .CreateField("Division", dbText)
            .Fields.Append .CreateField("District", dbText)
            .Fields.Append .CreateField("Location", dbText)
            .Fields.Append .CreateField("Customer Id", dbText)
            .Fields.Append .CreateField("Customer", dbText)
            .Fields.Append .CreateField("Phone", dbText)
            .Fields.Append .CreateField("Email", dbText)
            .Fields.Append .CreateField("Agreement Id", dbText)
            .Fields.Append .CreateField("Description", dbText)
            .Fields.Append .CreateField("Agreed Date", dbText)
            .Fields.Append .CreateField("Age (days)", dbInteger)
            .Fields.Append .CreateField("Cancelled By", dbText)
            .Fields.Append .CreateField("Cancelled Reason", dbText)
            .Fields.Append .CreateField("Last Due Date", dbDate)
            .Fields.Append .CreateField("Monthly Draft", dbCurrency)
            .Fields.Append .CreateField("Total Agreement Value", dbCurrency)
            .Fields.Append .CreateField("Total Refunded", dbCurrency)
            .Fields.Append .CreateField("Net", dbCurrency)
            .Fields.Append .CreateField("Referred By", dbText)
        End With
        
        db.TableDefs.Append tdfNew
        Set tdfNew = Nothing
        Set tdfNew = db.CreateTableDef("Member Export")
        
                
        With tdfNew
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
            .Fields.Append .CreateField("Gender", dbText)
            .Fields.Append .CreateField("DOB", dbDate)
            .Fields.Append .CreateField("Address", dbText)
            .Fields.Append .CreateField("City", dbText)
            .Fields.Append .CreateField("CA", dbText)
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
        
        db.TableDefs.Append tdfNew
        Set tdfNew = Nothing
        Set tdfNew = db.CreateTableDef("Attendance")
        
        With tdfNew
            .Fields.Append .CreateField("Club Name", dbText)
            .Fields.Append .CreateField("UserID", dbText)
            .Fields.Append .CreateField("Last Name", dbText)
            .Fields.Append .CreateField("First Name", dbText)
            .Fields.Append .CreateField("Check In Date / Time", dbDate)
            .Fields.Append .CreateField("Home Club", dbText)
            .Fields.Append .CreateField("Allowed", dbBoolean)
            .Fields.Append .CreateField("Denied", dbBoolean)
            .Fields.Append .CreateField("Gender", dbText)
            .Fields.Append .CreateField("DOB", dbDate)
            .Fields.Append .CreateField("Membership Type", dbText)
            .Fields.Append .CreateField("Prospect Type", dbText)
            .Fields.Append .CreateField("Responsible Person Last Name", dbText)
            .Fields.Append .CreateField("Responsible Person First Name", dbText)
            .Fields.Append .CreateField("Custom Field", dbText)
        End With
    
    db.TableDefs.Append tdfNew
    Set tdfNew = Nothing
        
    Set tdfNew = db.TableDefs("Attendance")
    
    With tdfNew
        .Fields("Allowed").Properties.Append .Fields("Allowed").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("Denied").Properties.Append .Fields("Denied").CreateProperty("DisplayControl", dbInteger, 106)
    End With
    
    Set tdfNew = Nothing
    
    On Error GoTo 0
    CreateDatabase = True
    
Handler:
    Debug.Print (Err.Description)
    On Error GoTo 0
    CreateDatabase = False
    
End Function
