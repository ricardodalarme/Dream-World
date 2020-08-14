VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12000
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5670
      TabIndex        =   2
      Top             =   240
      Width           =   570
   End
   Begin VB.Label lblTip 
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
      Height          =   210
      Left            =   2400
      TabIndex        =   1
      Top             =   7080
      Width           =   7155
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3720
      TabIndex        =   0
      Top             =   8640
      Width           =   4200
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Move localization
Public MoveX As Long
Public MoveY As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If KeyAscii = vbKeyEscape Then
        Call DestroyTCP
        frmLoad.Hide
        frmMenu.Show
        frmMenu.picLogin.Visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set form caption
    Me.Caption = Options.Game_Name & " (esc to cancel)"
    lblCaption.Caption = Options.Game_Name & " (esc to cancel)"
    
    ' Set the tip
    i = Rand(1, Max_Tips)
    If Len(Trim$(Tip(i))) > 0 Then lblTip.Caption = Trim$(Tip(i))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_MouseDown", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
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

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' When the form close button is pressed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Call DestroyTCP
        frmLoad.Hide
        frmMenu.Show
        frmMenu.picLogin.Visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

