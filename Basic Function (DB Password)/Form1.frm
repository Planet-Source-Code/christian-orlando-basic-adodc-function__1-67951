VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "campo1"
      DataSource      =   "Adodc"
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   3255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":0000
         Height          =   1815
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc 
         Height          =   375
         Left            =   1560
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Field 2"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Field 1"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function DatabaseConnect(AdoControl As Adodc)

AdoControl.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\prueba.mdb") & ";Persist Security Info=False;Jet OLEDB:Database Password=codj")
AdoControl.RecordSource = "Select * From tabla"
AdoControl.Refresh
End Function

Private Sub Command1_Click()
Adodc.Recordset.MoveFirst


Adodc.Recordset.Find "campo1 =" & "'" & Text1 & "'"

If Len(Text1) > 0 And Len(Text2) > 0 Then
    If Len(Text3) = 0 Then
        Adodc.Recordset.AddNew
        Adodc.Recordset.Fields(0) = Text1.Text
        Adodc.Recordset.Fields(1) = Text2.Text
        Call limpia
        MsgBox "The Data Has Been Saved", vbInformation, "Test"
        
    Else
        mb = MsgBox("The Data Already Exist.", vbCritical, "Error")
        Call limpia
    End If
Else
    mb = MsgBox("You Must Fill All The Fields.", vbCritical, "Error")
    
End If
End Sub

Private Sub Command2_Click()
searcx = InputBox("Enter The field1", "Searching")
book = Adodc.Recordset.Bookmark

Adodc.Recordset.MoveFirst
Do Until Adodc.Recordset.EOF Or Found
    If Adodc.Recordset.Fields("campo1") Like searcx Then
        Found = True
        Text1.Text = Adodc.Recordset.Fields(0)
        Text2.Text = Adodc.Recordset.Fields(1)
        Command1.Enabled = False
        Command3.Enabled = True
        Command4.Enabled = True

    Else
        Adodc.Recordset.MoveNext
    End If
Loop
    
If Found = False Then
    MsgBox "Record not found", vbInformation, "Test"
    Adodc.Recordset.Bookmark = book
    Command1.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = False

End If
End Sub

Private Sub Command3_Click()
Adodc.Recordset.Update
Adodc.Recordset.Fields(0) = Text1.Text
Adodc.Recordset.Fields(1) = Text2.Text
Call limpia
Command3.Enabled = False
Command1.Enabled = True
Command4.Enabled = False

End Sub

Private Sub Command4_Click()
Adodc.Recordset.Delete
Call limpia
Command4.Enabled = False
Command1.Enabled = True
Command3.Enabled = False

End Sub

Private Sub Form_Load()
DatabaseConnect Adodc
Command3.Enabled = False
Command4.Enabled = False

End Sub

Private Sub limpia()
Text1.Text = ""
Text2.Text = ""
End Sub
