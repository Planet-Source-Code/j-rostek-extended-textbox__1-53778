VERSION 5.00
Object = "*\A..\Projekt2.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   6780
      TabIndex        =   8
      Top             =   540
      Width           =   435
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Currency"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Textbox"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Richtext"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Number"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Percent"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      Caption         =   "Yes/No Checked"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin Projekt2.Etextbox Etextbox1 
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   60
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EnforceCurrencyFormat(ptxtBox As TextBox, _
    pintKeyascii As Integer)
    
    'Purpose: To enforce single decimal poin
    '     t two-decimal places maximum.
    'Usage: Call from the KeyPress event of
    '     a text box.


    Select Case pintKeyascii
        Case 45
          If InStr(ptxtBox.Text, "-") > 0 Then 'decimal point already exists...
            pintKeyascii = 0
          End If
          If ptxtBox.SelStart <> 0 Then
            pintKeyascii = 0
          End If

        
        Case 48 To 58 'numbers must be handled specially...
                      'but only if there is already a decimal
          '     point in the string...
          If InStr(ptxtBox.Text, ",") > 0 Then
              If ptxtBox.SelStart = Len(ptxtBox.Text) Then 'inserting at End of string...
                  If InStr(ptxtBox.Text, ",") + 2 = Len(ptxtBox.Text) Then 'already two digits To right of Decimal point...
                      pintKeyascii = 0
                  End If
              End If
          End If
        Case Is = 44 'decimal point...
            If InStr(ptxtBox.Text, ",") > 0 Then 'decimal point already exists...
                If ptxtBox.SelLength = 0 Then 'the existing Decimal point is Not selected...
                    'eat the keystroke or a second decimail
                    '     point would appear...
                    pintKeyascii = 0
                Else
                    Select Case InStr(ptxtBox.Text, ",") 'determine where the Decimal point occurs...
                        Case ptxtBox.SelStart To (ptxtBox.SelStart + ptxtBox.SelLength)
                        'the decimal point is selected, so it wi
                        '     ll be replaced...
                        Case Else
                        pintKeyascii = 0
                    End Select
            End If
        
        
    Else
        If Len(ptxtBox.Text) - ptxtBox.SelLength > ptxtBox.SelStart + 2 Then 'make sure it's ok To insert a Decimal point here...
            pintKeyascii = 0
        End If
    End If
    Case Is = 8 'backspace is always ok...
    Case Else 'all other characters should be eaten...
    pintKeyascii = 0
End Select

End Sub


Private Sub Form_Load()
Me.Option1(0).Value = True

End Sub

Private Sub Option1_Click(Index As Integer)
    Me.Etextbox1.cCurrency = False
    Me.Etextbox1.cNumber = False
    Me.Etextbox1.cTextbox = False
    Me.Etextbox1.cRichEdit = False
    Me.Etextbox1.cDateBox = False
    Me.Etextbox1.cPercent = False
    Me.Etextbox1.cYesNo = False
    Me.Etextbox1.cCaption = "Control " & Option1(Index).Caption
    Me.Etextbox1.cSubQuery = ""
    Me.Etextbox1.cDBFile = ""

Select Case Index
  Case 0
    Me.Etextbox1.cCurrency = True
  Case 1
    Me.Etextbox1.cNumber = True
  Case 2
    Me.Etextbox1.cSubQuery = "select * from Wathever"
    Me.Etextbox1.cDBFile = "C:\TheDBFile.mdb"
    Me.Etextbox1.cDBBoundColumn = 3
    Me.Etextbox1.cTextbox = True
  Case 3
    Me.Etextbox1.cRichEdit = True
  Case 4
    Me.Etextbox1.cDateBox = True
  Case 5
    Me.Etextbox1.cPercent = True
  Case 6
    Me.Etextbox1.cYesNo = True

End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
EnforceCurrencyFormat Text1, KeyAscii
End Sub
