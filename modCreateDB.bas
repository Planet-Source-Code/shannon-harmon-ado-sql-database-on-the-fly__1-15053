Attribute VB_Name = "modCreateDB"
Option Explicit

Public Function CreateDB(strDBPath As String) As Boolean

  On Error GoTo PROC_ERR
  
  Dim objTbl As New Table
  Dim objCat As New ADOX.Catalog
  Dim objKey As New ADOX.Key
  
  Dim strConnection As String
  
  strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & strDBPath
  
  objCat.Create strConnection 'Create the database
  
  'Create our inventory table and columns for the example
  objTbl.Name = "Inventory"
  objTbl.Columns.Append "Sku", adVarWChar, 15
  objTbl.Columns.Append "Description", adVarWChar, 50
  objTbl.Columns.Append "Price", adCurrency
  objTbl.Columns.Append "Cost", adCurrency
  
  objTbl.Columns("Description").Attributes = adColNullable  'Set description nullable
  
  objCat.Tables.Append objTbl 'Append this table to the database
  
  'Create our primary key
  objKey.Name = "Sku"
  objKey.Type = adKeyPrimary
  objKey.Columns.Append "Sku"
  
  objTbl.Keys.Append objKey 'Append this key to the database

  CreateDB = True 'No errors return true
  
PROC_EXIT:
  On Error Resume Next
  'Remove our objects from memory
  Set objKey = Nothing
  Set objTbl = Nothing
  Set objCat = Nothing
  Exit Function
  
PROC_ERR:
  CreateDB = False  'Error return false
  Resume PROC_EXIT  'Resume exit
End Function


