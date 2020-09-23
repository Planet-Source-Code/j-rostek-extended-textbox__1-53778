VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmQuery 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows-Standard
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2700
      Top             =   2460
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   582
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   60
      Width           =   2115
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   60
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmQuery.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
            LCID            =   1031
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
            LCID            =   1031
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "nach"
      Height          =   255
      Left            =   3660
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "suche in"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblBoundColumn 
      Caption         =   "0"
      Height          =   255
      Left            =   5340
      TabIndex        =   1
      Top             =   4500
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid1_DblClick()
If lblBoundColumn.Caption = "" Then lblBoundColumn.Caption = "0"
  
DataGrid1.Col = lblBoundColumn.Caption
Me.Tag = DataGrid1.Text
Me.Hide
End Sub



Private Sub Form_Resize()
On Error Resume Next
  DataGrid1.Left = 0
'  DataGrid1.Top = 0
  DataGrid1.Width = Me.Width - 150
  DataGrid1.Height = (Me.Height - DataGrid1.Top) - 150
End Sub

Private Sub Text1_Change()
Adodc1.Refresh

mywert = ""
For g = 0 To DataGrid1.Columns.Count - 1
  mywert = mywert & "[" & DataGrid1.Columns(g).Caption & "]" & ","
Next
mywert = Left(mywert, Len(mywert) - 1)
Set frmQuery.DataGrid1.DataSource = Nothing
Set frmQuery.DataGrid1.DataSource = Adodc1
cSubQuery = "select " & mywert & " from Kunden where [" & "Kundensuchname" & "] LIKE '" & Text1.Text & "*'"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Dokumente und Einstellungen\Administrator\Eigene Dateien\Feisthauer\N3AC2000.mdb;Mode=Read;Persist Security Info=False"
Adodc1.RecordSource = CStr(cSubQuery)
Adodc1.Refresh


'DataGrid1.ClearFields
DataGrid1.ReBind
DataGrid1.Refresh





 
End Sub
