VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picCharacters 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   210
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   27
      Top             =   4500
      Visible         =   0   'False
      Width           =   4650
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1650
         ItemData        =   "frmMenu.frx":3332
         Left            =   120
         List            =   "frmMenu.frx":3334
         TabIndex        =   28
         Top             =   840
         Width           =   4290
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "here"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   5
         Left            =   2220
         TabIndex        =   32
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label lblCharCommand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label lblCharCommand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3480
         TabIndex        =   30
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label lblCharCommand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   2040
         TabIndex        =   29
         Top             =   2640
         Width           =   315
      End
   End
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   210
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   4
      Top             =   4500
      Visible         =   0   'False
      Width           =   4650
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   975
         TabIndex        =   20
         Top             =   735
         Width           =   2715
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   2085
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblCAccept 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   1950
         TabIndex        =   17
         Top             =   3720
         Width           =   765
      End
      Begin VB.Label lblCChange 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   1
         Left            =   1920
         TabIndex        =   16
         Top             =   2760
         Width           =   105
      End
      Begin VB.Label lblCChange 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   2
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   105
      End
      Begin VB.Label lblCGender 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   2130
         TabIndex        =   14
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   960
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   210
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   4650
      Begin VB.TextBox txtRCaptcha 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   1395
         TabIndex        =   24
         Top             =   2895
         Width           =   1755
      End
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1395
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   2430
         Width           =   2715
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1395
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   1965
         Width           =   2715
      End
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   1395
         TabIndex        =   21
         Top             =   1500
         Width           =   2715
      End
      Begin VB.Label lblFunction 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click Here"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   4
         Left            =   1320
         TabIndex        =   12
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label lblRAccept 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   1920
         TabIndex        =   11
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label lblCaptcha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3210
         TabIndex        =   6
         Top             =   2895
         Width           =   870
      End
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   210
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   310
      TabIndex        =   0
      Top             =   4500
      Width           =   4650
      Begin VB.TextBox txtLUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   1395
         TabIndex        =   19
         Top             =   1500
         Width           =   2715
      End
      Begin VB.TextBox txtLPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1395
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1965
         Width           =   2715
      End
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1365
         TabIndex        =   7
         Top             =   2325
         Width           =   195
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click Here"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   2
         Left            =   960
         TabIndex        =   9
         Top             =   3960
         Width           =   810
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sign Up"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   1
         Left            =   1905
         TabIndex        =   8
         Top             =   405
         Width           =   615
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Start"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   2880
         Width           =   1305
      End
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   0
      ScaleHeight     =   2880
      ScaleWidth      =   12000
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   12000
      Begin VB.Label lblFunction 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Index           =   3
         Left            =   11400
         TabIndex        =   10
         Top             =   2520
         Width           =   435
      End
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   15
      Width           =   630
   End
   Begin VB.Image imgButton 
      Height          =   180
      Index           =   1
      Left            =   11565
      Top             =   30
      Width           =   180
   End
   Begin VB.Image imgButton 
      Height          =   180
      Index           =   2
      Left            =   11775
      Top             =   30
      Width           =   180
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Move localization
Public MoveX As Long
Public MoveY As Long

' Captha
Private Sub Captcha(ByVal Label As Label)
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Clear label caption
    Label.Caption = vbNullString
    
    ' Set the label caption
    For i = 1 To 5
        Label.Caption = Label.Caption & Rand(0, 9)
    Next
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Captha", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' New Char
Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharSprite = 0
    newCharClass = cmbClass.ListIndex + 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' general menu stuff
    lblCaption.Caption = Options.Game_Name
    Me.Caption = Options.Game_Name

    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.SavePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.Value = Options.SavePass
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Move form
    If Button = 1 Then
        MoveX = x
        MoveY = y
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Move form
    If Button = 1 Then
        Me.Left = (Me.Left - MoveX) + x
        Me.top = (Me.top - MoveY) + y
    End If
    
    ' reset all buttons
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not EnteringGame Then DestroyGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index = 1 Then
        Me.WindowState = vbMinimized
    End If
    
    If Index = 2 Then
        DestroyGame
    End If
    
    ' play sound
    Play_Sound Button_Click
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    changeButtonState_Menu Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    If Not MenuButton(Index).State = 2 Then ' make sure we're not clicking
        changeButtonState_Menu Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Menu = Index Then
        Play_Sound Button_Hover
        LastButtonSound_Menu = Index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Menu -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MenuState(MENU_STATE_ADDCHAR)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCaptcha_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Captcha lblCaptcha
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCaptcha_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCChange_Click(Index As Integer)
    Dim SpriteCount As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            If newCharSex = SEX_MALE Then SpriteCount = UBound(Class(newCharClass).MaleSprite) Else SpriteCount = UBound(Class(newCharClass).FemaleSprite)
            If newCharSprite >= SpriteCount Then newCharSprite = 0 Else newCharSprite = newCharSprite + 1
        Case 2
            If newCharSex = SEX_MALE Then SpriteCount = UBound(Class(newCharClass).MaleSprite) Else SpriteCount = UBound(Class(newCharClass).FemaleSprite)
            If newCharSprite <= 0 Then newCharSprite = SpriteCount Else newCharSprite = newCharSprite - 1
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCChange_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCGender_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If newCharSex = SEX_MALE Then
        lblCGender.Caption = "Female"
        newCharSex = SEX_FEMALE
    Else
        lblCGender.Caption = "Male"
        newCharSex = SEX_MALE
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCGender_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCharCommand_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            If Len(Trim$(CharData(lstChars.ListIndex + 1).Name)) > 0 Then
                MsgBox "There is already a character in this slot!"
                Exit Sub
            End If
            
            Call MenuState(MENU_STATE_NEWCHAR)
        Case 2
            If Len(Trim$(CharData(lstChars.ListIndex + 1).Name)) <= 0 Then
                MsgBox "There is no character in this slot!"
                Exit Sub
            End If
        
            Call MenuState(MENU_STATE_USECHAR)
        Case 3
            If Len(Trim$(CharData(lstChars.ListIndex + 1).Name)) <= 0 Then
                MsgBox "There is no character in this slot!"
                Exit Sub
            End If

            If MsgBox("Are you sure you wish to delete this character?", vbYesNo, Options.Game_Name) = vbYes Then
                Call MenuState(MENU_STATE_DELCHAR)
            End If
    End Select

    Play_Sound Button_Click

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblFunction_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblFunction_Click(Index As Integer)
    Dim Name As String, Password As String, PasswordAgain As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            If Not picRegister.Visible Then
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = False
                picRegister.Visible = True
                picCharacter.Visible = False
                picCharacters.Visible = False
                Captcha lblCaptcha
            End If
        Case 2
            If Not picCredits.Visible Then
                DestroyTCP
                picCredits.Visible = True
                picLogin.Visible = False
                picRegister.Visible = False
                picCharacter.Visible = False
                picCharacters.Visible = False
            End If
        Case 3
            If Not picLogin.Visible Then
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = True
                picRegister.Visible = False
                picCharacter.Visible = False
                picCharacters.Visible = False
            End If
        Case 4
            If Not picLogin.Visible Then
                DestroyTCP
                picCredits.Visible = False
                picLogin.Visible = True
                picRegister.Visible = False
                picCharacter.Visible = False
                picCharacters.Visible = False
            End If
        Case 5
            DestroyTCP
            picCredits.Visible = False
            picLogin.Visible = True
            picRegister.Visible = False
            picCharacter.Visible = False
            picCharacters.Visible = False
    End Select

    Play_Sound Button_Click

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblFunction_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Register
Private Sub lblRAccept_Click()
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(Name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(Name) Then
            Call MsgBox("Invalid User.")
            Exit Sub
        End If
        
        If txtRCaptcha.text <> lblCaptcha.Caption Then
            Captcha lblCaptcha
            Call MsgBox("InColorrect Code.")
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
