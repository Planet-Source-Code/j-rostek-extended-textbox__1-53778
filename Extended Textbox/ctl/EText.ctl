VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Etextbox 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   PropertyPages   =   "EText.ctx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   5355
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   2040
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22675457
      CurrentDate     =   38051
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   2595
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"EText.ctx":0010
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1260
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   900
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   540
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   180
      Width           =   2655
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   315
      Left            =   60
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "Etextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim mCurrency As Boolean
Dim mNumber As Boolean
Dim mPercent As Boolean
Dim mTextbox As Boolean
Dim mDateBox As Boolean
Dim mRichEdit As Boolean
Dim mYesNo As Boolean
Dim mSubQuery As String
Dim mDBFile As String
Dim mCaption As String
Dim mDBBoundColumn
Dim ChooseDate
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Private Const LOCALE_SENGCOUNTRY = &H1002       '  English name of country
Private Const LOCALE_SCURRENCY = &H14           '  local monetary symbol
Private Const LOCALE_SDATE = &H1D               '  date separator
Private Const LOCALE_SLONGDATE = &H20           '  long date format string
Private Const LOCALE_SSHORTDATE = &H1F          '  short date format string
Private Const LOCALE_SDECIMAL = &HE             '  decimal separator
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
  X As Long                                     'x coordinate of the mouse
  Y As Long                                     'y coordinate of the mouse
End Type
Dim pp As POINTAPI
Public Function GetDecimalSeparator() As String
'###################################################
' Get the Decimal Separator
'###################################################
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, buffer, 99)
   GetDecimalSeparator = LPSTRToVBString(buffer)
End Function
Private Function GetCurrencySymbol() As String
'###################################################
' Get the Currency Symbol
'###################################################
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, buffer, 99)
   GetCurrencySymbol = LPSTRToVBString(buffer)
End Function
Private Function LPSTRToVBString$(ByVal s$)
'###################################################
' Dirty little Helper
'###################################################
   Dim nullpos&
   nullpos& = InStr(s$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function
Private Function GetDateSeparator() As String
'###################################################
' Date Seperator
'###################################################
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, buffer, 99)
   GetDateSeparator = LPSTRToVBString(buffer)
End Function
Private Function GetLongDateFormat() As String
'###################################################
' Get Long Date Format
'###################################################
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, buffer, 99)
   GetLongDateFormat = LPSTRToVBString(buffer)
End Function

Private Sub Check1_GotFocus()
Shape1.Visible = True
End Sub

Private Sub Check1_LostFocus()
Shape1.Visible = False
End Sub

Private Sub Command1_Click()
If Dir(cDBFile) = "" Then Exit Sub
frmQuery.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cDBFile & ";Persist Security Info=False "
frmQuery.Adodc1.RecordSource = cSubQuery
frmQuery.Adodc1.Refresh
frmQuery.DataGrid1.ClearFields
frmQuery.DataGrid1.ReBind
frmQuery.lblBoundColumn.Caption = mDBBoundColumn

frmQuery.Combo1.Clear
For g = 0 To frmQuery.DataGrid1.Columns.Count - 1
  frmQuery.Combo1.AddItem frmQuery.DataGrid1.Columns(g).Caption
Next

frmQuery.Combo1.Tag = cDBFile
frmQuery.Text1.Tag = cSubQuery
frmQuery.Show 1
If Text1.Visible = True Then Text1.Text = frmQuery.Tag
If Text2.Visible = True Then Text2.Text = frmQuery.Tag
If Text3.Visible = True Then Text3.Text = frmQuery.Tag
If Text4.Visible = True Then Text4.Text = frmQuery.Tag
If RichTextBox1.Visible = True Then RichTextBox1.Text = frmQuery.Tag
Unload frmQuery
End Sub

Private Sub DTPicker1_GotFocus()
Shape1.Visible = True
End Sub

Private Sub DTPicker1_LostFocus()
Shape1.Visible = False
End Sub

Private Sub RichTextBox1_GotFocus()
Shape1.Visible = True
End Sub

Private Sub RichTextBox1_LostFocus()
Shape1.Visible = False
End Sub

Private Sub Text1_GotFocus()
Shape1.Visible = True
'###################################################
' Format the Currency
'###################################################
If Text1.Text <> "" Then
Text1.Text = Format(Text1.Text, "##,##0.00")
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
EnforceCurrencyFormat Text1, KeyAscii
End Sub

Private Sub Text1_LostFocus()
Shape1.Visible = False
'###################################################
' add the Currency Symbol to text1
'###################################################
Text1.Text = Format(Text1.Text, "##,##0.00")
Text1.Text = Text1.Text & " " & GetCurrencySymbol
End Sub

Private Sub Text2_GotFocus()
Shape1.Visible = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
EnforceNumberFormat Text2, KeyAscii
End Sub

Private Sub Text2_LostFocus()
Shape1.Visible = False
End Sub


Private Sub Text3_GotFocus()
Shape1.Visible = True
'###################################################
' replace the percent Symbol from text3
'###################################################
Text3.Text = Replace(Text3.Text, " %", "")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
EnforceCurrencyFormat Text3, KeyAscii
End Sub

Private Sub Text3_LostFocus()
Shape1.Visible = False
'###################################################
' add the percent Symbol to text3
'###################################################
Text3.Text = Format(Text3.Text, "##,##0.00")
Text3.Text = Text3.Text & " %"
End Sub

Private Sub Text4_GotFocus()
Shape1.Visible = True
End Sub

Private Sub Text4_LostFocus()
Shape1.Visible = False
End Sub

Private Sub UserControl_Initialize()
Shape1.Visible = False
End Sub

Private Sub UserControl_Resize()
'###################################################
' Size all controls to the correct position
'###################################################
lblCaption.Top = 20
lblCaption.Left = 20
tempTop = lblCaption.Top + lblCaption.Height + 20
Shape1.Left = 0
Shape1.Top = tempTop
Shape1.Width = UserControl.Width
Shape1.Height = (UserControl.Height) - tempTop
Text1.Left = 20
Text1.Top = tempTop + 20
Text1.Width = UserControl.Width - 30
Text1.Height = (UserControl.Height - 30) - tempTop
Text2.Left = 20
Text2.Top = tempTop + 20
Text2.Width = UserControl.Width - 30
Text2.Height = (UserControl.Height - 30) - tempTop
Text3.Left = 20
Text3.Top = tempTop + 20
Text3.Width = UserControl.Width - 30
Text3.Height = (UserControl.Height - 30) - tempTop
Text4.Left = 20
Text4.Top = tempTop + 20
Text4.Width = UserControl.Width - 30
Text4.Height = (UserControl.Height - 30) - tempTop

RichTextBox1.Left = 20
RichTextBox1.Top = tempTop + 20
RichTextBox1.Width = UserControl.Width - 30
RichTextBox1.Height = (UserControl.Height - 30) - tempTop


Check1.Left = 20
Check1.Top = tempTop + 20
Check1.Width = UserControl.Width - 30
Check1.Height = (UserControl.Height - 30) - tempTop



Command1.Top = tempTop + 20
Command1.Width = UserControl.Height - 30
Command1.Height = (UserControl.Height - 30) - tempTop
Command1.Left = (UserControl.Width - Command1.Width) - 20

DTPicker1.Left = 20
DTPicker1.Top = tempTop + 20
DTPicker1.Width = UserControl.Width - 30
DTPicker1.Height = (UserControl.Height - 30) - tempTop

End Sub

Private Sub EnforceCurrencyFormat(ptxtBox As TextBox, pintKeyascii As Integer)
Select Case pintKeyascii
  Case 45
    If InStr(ptxtBox.Text, "-") > 0 Then                                          'minus already exists...
      pintKeyascii = 0
    End If
    If ptxtBox.SelStart <> 0 Then
      pintKeyascii = 0
    End If
  Case 48 To 58                                                                   'numbers
    If InStr(ptxtBox.Text, GetDecimalSeparator) > 0 Then
      If ptxtBox.SelStart = Len(ptxtBox.Text) Then                                'inserting at End of string...
        If InStr(ptxtBox.Text, GetDecimalSeparator) + 2 = Len(ptxtBox.Text) Then  'already two digits To right of GetDecimalSeparator
          pintKeyascii = 0
        End If
      End If
    End If
  Case Is = 44                                                                    'Decimal Separator
    If InStr(ptxtBox.Text, GetDecimalSeparator) > 0 Then                          'Decimal Separator already exists
      If ptxtBox.SelLength = 0 Then                                               'the existing Decimal Separator is Not selected
        pintKeyascii = 0
      Else
        Select Case InStr(ptxtBox.Text, GetDecimalSeparator)                      'determine where the Decimal Separator occurs
          Case ptxtBox.SelStart To (ptxtBox.SelStart + ptxtBox.SelLength)
                                                                                  'the Decimal Separator is selected, so it will be replaced
          Case Else
          pintKeyascii = 0
        End Select
      End If
    Else
      If Len(ptxtBox.Text) - ptxtBox.SelLength > ptxtBox.SelStart + 2 Then        'make sure it's ok To insert a Decimal Separator point here...
          pintKeyascii = 0
      End If
    End If
  Case Is = 8                                                                     'backspace is always ok...
  Case Else                                                                       'all other characters should be eaten...
  pintKeyascii = 0
End Select
End Sub

Public Property Get cCurrency() As Boolean
Attribute cCurrency.VB_ProcData.VB_Invoke_Property = "Propertys"
  cCurrency = mCurrency
End Property

Public Property Let cCurrency(ByVal vNewValue As Boolean)
  mCurrency = vNewValue
  SetControls
End Property
Public Property Get cNumber() As Boolean
Attribute cNumber.VB_ProcData.VB_Invoke_Property = "Propertys"
  cNumber = mNumber
End Property

Public Property Let cNumber(ByVal vNewValue As Boolean)
  mNumber = vNewValue
  SetControls
End Property

Public Property Get cPercent() As Boolean
Attribute cPercent.VB_ProcData.VB_Invoke_Property = "Propertys"
  cPercent = mPercent
End Property

Public Property Let cPercent(ByVal vNewValue As Boolean)
  mPercent = vNewValue
  SetControls
End Property
Public Property Get cTextbox() As Boolean
Attribute cTextbox.VB_ProcData.VB_Invoke_Property = "Propertys"
  cTextbox = mTextbox
End Property

Public Property Let cTextbox(ByVal vNewValue As Boolean)
  mTextbox = vNewValue
  SetControls
End Property

Public Property Get cDateBox() As Boolean
Attribute cDateBox.VB_ProcData.VB_Invoke_Property = "Propertys"
  cDateBox = mDateBox
End Property

Public Property Let cDateBox(ByVal vNewValue As Boolean)
  mDateBox = vNewValue
  SetControls
End Property


Public Property Get cRichEdit() As Boolean
  cRichEdit = mRichEdit
End Property

Public Property Let cRichEdit(ByVal vNewValue As Boolean)
  mRichEdit = vNewValue
  SetControls
End Property

Public Property Get cYesNo() As Boolean
  cYesNo = mYesNo
End Property

Public Property Let cYesNo(ByVal vNewValue As Boolean)
  mYesNo = vNewValue
  SetControls
End Property
Public Property Get cSubQuery() As String
  cSubQuery = mSubQuery
End Property

Public Property Let cSubQuery(ByVal vNewValue As String)
  mSubQuery = vNewValue
End Property
Public Property Get cDBFile() As String
  cDBFile = mDBFile
End Property

Public Property Let cDBFile(ByVal vNewValue As String)
  mDBFile = vNewValue
End Property

Public Property Get cDBBoundColumn() As String
  cDBBoundColumn = mDBBoundColumn
End Property

Public Property Let cDBBoundColumn(ByVal vNewValue As String)
  mDBBoundColumn = vNewValue
End Property

Public Property Get cCaption() As String
  cCaption = mCaption
End Property

Public Property Let cCaption(ByVal vNewValue As String)
  mCaption = vNewValue
  lblCaption = vNewValue
End Property

Private Sub SetControls()
'##############################################
' Set the current visible Control
'##############################################
Command1.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
RichTextBox1.Visible = False
Check1.Visible = False
DTPicker1.Visible = False
  If mCurrency = True Then
    Text1.Visible = True
  End If
  If mNumber = True Then
    Text2.Visible = True
  End If
  If mPercent = True Then
    Text3.Visible = True
  End If
  If mTextbox = True Then
    Text4.Visible = True
  End If
  If mRichEdit = True Then
    RichTextBox1.Visible = True
  End If
  If mDateBox = True Then
    DTPicker1.Visible = True
  End If
  If mYesNo = True Then
    Check1.Visible = True
  End If
  If mSubQuery <> "" And cDBFile <> "" Then
    Command1.Visible = True
  End If
End Sub

Private Sub EnforceNumberFormat(ptxtBox As TextBox, pintKeyascii As Integer)
    Select Case pintKeyascii
      Case 48 To 58 'numbers must be handled specially...
      Case Is = 8   'backspace is always ok...
      Case Else     'all other characters should be eaten...
      pintKeyascii = 0
    End Select
End Sub

Private Sub EnforceDateFormat(ptxtBox As TextBox, pintKeyascii As Integer)
    pintKeyascii = 0
End Sub



