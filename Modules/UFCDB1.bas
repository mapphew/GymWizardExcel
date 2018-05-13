Attribute VB_Name = "UFCDB1"
Sub CreateDatabase()

    ' ' Dim
    Dim Catalog As ADOX.Catalog
    Dim FSO As New FileSystemObject
    
    ' ' Create DB
    If Not FSO.FileExists(Range("MAINDIR").Value & "\UFCDB.accdb") And Range("MAINDIR").Value <> "" Then
        Set Catalog = New ADOX.Catalog
        Catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path & "\Gym Wizard\UFCDB.accdb"
    End If
    
    Range("DBDIR").Value = Range("MAINDIR").Value & "\UFCDB.accdb"
    
 
End Sub
