VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Insert Query Builder Example"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFunctions 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   405
      Index           =   1
      Left            =   5085
      TabIndex        =   9
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton cmdFunctions 
      Caption         =   "&Add To DB"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   3705
      TabIndex        =   8
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   825
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "0.00"
      Text            =   "0.00"
      Top             =   1530
      Width           =   870
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   825
      MaxLength       =   7
      TabIndex        =   2
      Tag             =   "0.00"
      Text            =   "0.00"
      Top             =   1095
      Width           =   870
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   1
      Left            =   825
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   660
      Width           =   5415
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   825
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   225
      Width           =   1740
   End
   Begin VB.Line lineSplitter 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6420
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line lineSplitter 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   15
      X2              =   6435
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&COST:"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   7
      Top             =   1575
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&PRICE:"
      Height          =   195
      Index           =   2
      Left            =   225
      TabIndex        =   6
      Top             =   1140
      Width           =   525
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&DESC:"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   705
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "&SKU:"
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   4
      Top             =   270
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Notes:
'
'This example was not meant to help you learn how to get data or update data
'from a database.
'
'This project references ADO 2.5 and ADO Ext 2.5 DLL without those of course
'it will not work, though you may update them to newer versions as needed.
'
'Feel free to use any of this code in your application but please give credit where
'credit is due.  If you have any questions or find something I screwed up, please
'feel free to email me: sharmonvpc@zdnetonebox.com
'
'Thanks - Shannon Harmon
'
'Originally decided to make this just for the modSQLBuilder code, then when
'making the example for you to use, I did everything else for some unknown
'reason.
'
'
'Use this code at your discretion, there is no implied warranty and use is at
'your own risk, given as freeware on 2/6/2001
'
'I am sure there are better ways to do things or even errors with my code,
'I wrote it in about an hour or two so if you find a problem and want to contribute
'to the code email me at sharmonvpc@zdnetonebox.com, if you just want to be rude,
'my email address is norudeemails@rude.com :)
'
Option Explicit

Dim arrColNames(3) As String
Dim arrColValues(3) As String

Dim objConn As New ADODB.Connection

Private Sub cmdFunctions_Click(Index As Integer)
  
  On Error Resume Next
  
  Select Case Index
    Case 0  'Add
      If Trim(txtField(0).Text) = "" Then 'Generic quick validation
        MsgBox "You must enter a value for SKU", vbInformation, "Insert aborted"
      Else  'Passed validation, allow record to be inserted
        InsertRecord
      End If
      
      txtField(0).SetFocus
      
    Case 1  'Exit
      Unload Me
  End Select
End Sub

Private Sub Form_Load()
  
  On Error GoTo PROC_ERR
  
  Dim strDBPath As String
  
  strDBPath = App.Path
  If Right(strDBPath, 1) <> "\" Then strDBPath = strDBPath & "\"
  strDBPath = strDBPath & "example.mdb"
  
  'Create the database if it is not yet been created
  If Trim(Dir(strDBPath)) = "" Then
    If Not CreateDB(strDBPath) Then 'Error creating database
      If Trim(Dir(strDBPath)) <> "" Then Kill strDBPath 'Delete if partially made
      
      MsgBox "Error creating database, shutting down", vbCritical, "Critical Error"
      Unload Me
      
      Exit Sub
    End If
  End If
      
  'Set connection properties and open
  With objConn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & strDBPath
    .Open , "Admin", ""
  End With
  
  ClearFields 'Set our textbox values to default new record
  
  'Set our array for the column names
  arrColNames(0) = "Sku"
  arrColNames(1) = "Description"
  arrColNames(2) = "Price"
  arrColNames(3) = "Cost"
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & vbCrLf & _
         "Description: " & Err.Description, vbCritical, "An error has occured"
         
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next
  
  If objConn.State = adStateOpen Then objConn.Close
  Set objConn = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
  
  txtField(Index).SelStart = 0
  txtField(Index).SelLength = Len(txtField(Index).Text)
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
  'Generic conversions and validations
  
  Select Case Index 'Conversions
    Case 0  'Sku
      KeyAscii = Asc(UCase(Chr(KeyAscii)))  'Convert to all caps
  
  End Select
  
  Select Case Index 'Validations
    Case 0  'Sku
      If KeyAscii = 32 Then KeyAscii = 0: Exit Sub  'No spaces
  
    Case 2, 3 'Price, Cost
      If KeyAscii = 8 Then Exit Sub
      If KeyAscii = 32 Then KeyAscii = 0: Exit Sub  'No spaces
      If KeyAscii = 45 Then KeyAscii = 0: Exit Sub  'No negatives
      
      Dim strFull As String
      strFull = txtField(Index) & Chr(KeyAscii)
      If Not IsNumeric(strFull) Then KeyAscii = 0: Exit Sub
  
  End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
  'Generic format of money fields
  
  If Index = 2 Or Index = 3 Then
    If Trim(txtField(Index)) = "" Then txtField(Index) = "0.00"
    txtField(Index) = Round(txtField(Index), 2)
    txtField(Index) = Format(txtField(Index), "0.00")
  End If
End Sub

Private Sub ClearFields()
  
  Dim objTextbox As TextBox
  
  'Notice I am using the textbox tag value
  'as the default value for the field on clear
  
  For Each objTextbox In txtField
    objTextbox.Text = objTextbox.Tag
  Next objTextbox
End Sub

Private Sub InsertRecord()

  On Error GoTo PROC_ERR
  
  Dim i As Integer
  
  For i = 0 To UBound(arrColValues)
    arrColValues(i) = txtField(i).Text
  Next i
  
  Dim strSQL As String
  strSQL = BuildInsertSQL("Inventory", objConn, arrColNames, arrColValues)
  
  Debug.Print strSQL  'Look in your debug window here is the created query!
  
  objConn.Execute strSQL  'Adds the record to the database.
  
  ClearFields

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & vbCrLf & _
         "Description: " & Err.Description, vbCritical, "An error has occured"
  Resume PROC_EXIT
End Sub

