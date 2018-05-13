VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UFCDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pDatabase As DAO.Database
Private pDatabasePath As String

Private Sub CreateDB(ByVal Path As String)

    ' ' DIM
    Dim Catalog As New ADOX.Catalog
    
    ' ' Create
    Catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path
    
End Sub

Private Sub CreateCPDTable()
    
    ' ' Create New Table Def
    Dim TDF As DAO.TableDef
    Set TDF = pDatabase.CreateTableDef("Club Past Due")
    
    ' ' Set Fields
    With TDF
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
    pDatabase.TableDefs.Append TDF
    Set TDF = Nothing
    
    ' ' Clean it up
    Set TDF = pDatabase.TableDefs("Club Past Due")
    With TDF
        .Fields("Has Payment Details on File?").Properties.Append .Fields("Has Payment Details on File?").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("In Collections Status?").Properties.Append .Fields("In Collections Status?").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("Is Return?").Properties.Append .Fields("Is Return?").CreateProperty("DisplayControl", dbInteger, 106)
    End With
    
    ' ' set nothing
    Set TDF = Nothing
    
End Sub
    
Private Sub CreateICDTable()
    ' ' Create New Table Def
    Dim TDF As DAO.TableDef
    Set TDF = pDatabase.CreateTableDef("Invoices Coming Due")
    
    With TDF
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
    
    pDatabase.TableDefs.Append TDF
    Set TDF = pDatabase.TableDefs("Invoices Coming Due")
    
    With TDF
        .Fields("Has CC").Properties.Append .Fields("Has CC").CreateProperty("DisplayControl", dbInteger, 106)
        .Fields("Has ACH").Properties.Append .Fields("Has ACH").CreateProperty("DisplayControl", dbInteger, 106)
    End With
    
End Sub

Private Sub CheckDB()

    Dim Tables As New ArrayList
    
    Dim table As DAO.TableDef
    
    For Each table In pDatabase.TableDefs
        Tables.Add (table.Name)
        Debug.Print (table.Name)
    Next
    
    If Not IsInArrayList("Club Past Due", Tables) Then
        CreateCPDTable
    End If
    
    If Not IsInArrayList("Invoices Coming Due", Tables) Then
        CreateICDTable
    End If
End Sub
    
Private Sub Class_Initialize()
    ' ' Dim
    Dim FSO As New FileSystemObject
    
    ' ' Set Path
    pDatabasePath = Range("DBDIR").Value
    
    If pDatabasePath <> "" Then
        If Not FSO.FileExists(pDatabasePath) Then
            CreateDB pDatabasePath
        End If
    Else:
        pDatabasePath = Range("MAINDIR").Value & "\UFCDB.accdb"
        
        If pDatabasePath <> "" And Not FSO.FileExists(pDatabasePath) Then
            CreateDB pDatabasePath
        End If
        
        Range("DBDIR").Value = pDatabasePath
    End If
    
    Set pDatabase = OpenDatabase(pDatabasePath)
    
    CheckDB
End Sub


Private Sub Class_Terminate()
    pDatabase.Close
    
    
End Sub
