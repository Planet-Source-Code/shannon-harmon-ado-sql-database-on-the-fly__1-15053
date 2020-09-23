Attribute VB_Name = "modSQLBuilder"
Option Explicit

'*************************************************************
'Procedure:     Public Method BuildInsertSQL
'Created on:    02/05/01
'Created by:    Shannon Harmon
'Description:   Builds an insert SQL and formats field values.
'               Checks if column is nullable and sets query
'               appropriately when needed.  Trims all values.
'
'Parameters:
'strTableName   Table name to build query for
'objConn        Open connection object to use
'arrColNames()  Array containing column names - Starting at 0
'arrColValues() Array containing column values - Starting at 0
'
'Notice:        Pass all values in the arrays as strings even
'               if they are numeric values, the function will
'               properly format them.  Converts all empty
'               values to NULL if field can be null, if a
'               field is a numeric which can't be null and is
'               an empty value, this function sets it to 0.
'               There is no error handling for this function
'               because the only time it will have an error is
'               when you send it invalid data, which needs to
'               be caught during design time.  Be sure to have
'               a value for each column, even if it's empty.
'               Does not test your values for valid data such
'               as date values, etc... You should validate
'               your data before trying to do an insert or
'               restrict invalid data if possible.
'
'Requires:      Reference to ADO Ext 2.5 DLL or higher.
'Tested with:   VB6 SP4 With ADO 2.5 & 2.6, SQL Server 7
'               Should work with any ADO compatible database
'               such as MS Access.
'*************************************************************
Public Function BuildInsertSQL(strTableName As String, _
                               objConn As ADODB.Connection, _
                               arrColNames() As String, _
                               arrColValues() As String) As String

  Dim objCat As New ADOX.Catalog                      'Catalog object
  Dim objTable As New ADOX.Table                      'Table object
  Dim objColumn As New ADOX.Column                    'Column object
  
  Set objCat.ActiveConnection = objConn               'Set Catalog connection object
  Set objTable = objCat.Tables(strTableName)          'Set Current table
  
  Dim i As Integer                                    'For/Next loop variable
  Dim intFields As Integer                            'Total number of fields (starting at 0)
  Dim strSQL As String                                'String to build SQL on temporarily
  
  intFields = UBound(arrColNames)                     'Get total number of columns (starting at 0)
  
  strSQL = "INSERT INTO " & strTableName & " ("       'Start INSERT query
  
  For i = 0 To intFields                              'Get all column names for query
    strSQL = strSQL & "[" & arrColNames(i) & "]"      'Add brackets to name in case of space in name
    
    If i <> intFields Then                            'If not at the last column then add ", "
      strSQL = strSQL & ", "
    Else                                              'At last column start VALUES query
      strSQL = strSQL & ") VALUES ("
    End If
  Next i
  
  
  For i = 0 To intFields                              'Get all field values for query
    Set objColumn = objTable.Columns(arrColNames(i))  'Set our column object to current column
    
    If (objColumn.Attributes = adColNullable) And _
    Trim(arrColValues(i)) = "" Then                   'If it's nullable and empty set to null
      
      strSQL = strSQL & "NULL"
    
    ElseIf IsColNumeric(objColumn.Type) Then          'If it's numeric format column
      
      If Trim(arrColValues(i)) = "" Then              'Empty numeric, set to 0 by default
        strSQL = strSQL & "0"
      Else                                            'Not empty numeric, set to value
        strSQL = strSQL & arrColValues(i)
      End If
    
    Else                                              'It's a string value, add single quotes
                                                      'and convert single quotes in value to double quotes
      strSQL = strSQL & "'" & _
      Trim(Replace(arrColValues(i), "'", "''")) & "'"
    
    End If
    
    If i <> intFields Then                            'If not at the last column then add ", "
      strSQL = strSQL & ", "
    Else                                              'Add last column close query
      strSQL = strSQL & ")"
    End If
  Next i
    
  BuildInsertSQL = strSQL                             'Return our query
    
  Set objColumn = Nothing                             'Release objects from memory
  Set objTable = Nothing
  Set objCat = Nothing
End Function

'*************************************************************
'Procedure:     Public Method IsFieldNumeric
'Created on:    02/05/01
'Created by:    Shannon Harmon
'Description:   Returns whether an ado column type is numeric
'
'Parameters:
'lngType        Ado column type property value
'
'Notice:        I think i got all the field types correct,
'               you may want to double check me, it's late:)
'
'Requires:      Reference to ADO Ext 2.5 DLL or higher.
'Tested with:   VB6 SP4 With ADO 2.5 & 2.6, SQL Server 7
'               Should work with any ADO compatible database
'               such as MS Access.
'*************************************************************
Public Function IsColNumeric(lngType As ADOX.DataTypeEnum) As Boolean
  Select Case lngType
    Case adBigInt, adBinary, adBoolean, adChapter, adCurrency, _
         adDecimal, adDouble, adError, adFileTime, adGUID, adInteger, _
         adLongVarBinary, adNumeric, adSingle, adSmallInt, adTinyInt, _
         adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, _
         adVarBinary, adVarNumeric
         
      IsColNumeric = True
    
    Case Else
      
      IsColNumeric = False
  End Select
End Function


