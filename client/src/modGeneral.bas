Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' randomize rnd's seed
    Randomize
    
    ' Load tips
    LoadTips
    
    ' set loading screen
    loadGUI True
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions
    
    ' load gui
    Call SetStatus("Loading interface...")
    loadGUI
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "spells"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "titles"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "panoramas"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "questbuttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"

    ' initialize DirectX
    Call SetStatus("Initializing DirectX...")
    InitDirectDraw
    
    ' load music/sound engine
    Init_Music
    
    ' player menu music
    Play_Music Menu_Music
    
    ' Reset values
    Ping = -1
    
    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    
    ' Init TCP
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    
    ' Clear game values
    Call SetStatus("Clearing game values...")
    Call ClearGameData

    ' Set game values
    Call SetStatus("Setting game values...")

    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' clear
    GettingMap = True

    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = True
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' hide the load form
    frmLoad.Visible = False

    ' Update the form with the game's name before it's loaded
    Call SetStatus("Loading Menu...")
    frmMain.lblCaption.Caption = Options.Game_Name
    frmMain.Caption = Options.Game_Name
    frmMenu.Visible = True
    
    ' Init menu loop
    InMenu = True
    MenuLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub loadGUI(Optional ByVal loadingScreen As Boolean = False)
    Dim i As Long

    ' if we can't find the interface
    On Error GoTo errorhandler
    
    ' loading screen
    If loadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If

    ' menu
    frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    frmMenu.picCharacters.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\characterS.jpg")
    ' main
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main.jpg")
    frmMain.picInventory.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\inventory.jpg")
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    frmMain.picSpells.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\skills.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\options.jpg")
    frmMain.picParty.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_spell.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")
    frmMain.picQuest.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\Quest.jpg")
    frmMain.picTitleDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_title.jpg")
    frmMain.picTempTitle.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTitles.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\titles.jpg")
    ' main - bars
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.jpg")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    ' main - party bars
    For i = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(i).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_spirit.jpg")
    Next
    
    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.Width
    SPRBar_Width = frmMain.imgMPBar.Width
    EXPBar_Width = frmMain.imgEXPBar.Width
    ' party
    Party_HPWidth = frmMain.imgPartyHealth(1).Width
    Party_SPRWidth = frmMain.imgPartySpirit(1).Width
    
    Exit Sub
    
    ' let them know we can't load the GUI
errorhandler:
    MsgBox "Cannot find one or more interface images." & vbNewLine & "If they exist then you have not extracted the project properly." & vbNewLine & "Please follow the installation instructions fully.", vbCritical
    DestroyGame
    Exit Sub
End Sub

Public Sub MenuState(ByVal State As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Hide load form
    frmLoad.Visible = True
    
    ' Close windows
    frmMenu.Visible = False
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picCharacters.Visible = False

    ' Checks if it is possible to connect to the server
    If Not ConnectToServer Then
        frmMenu.picLogin.Visible = True
        frmLoad.Visible = False
        frmMenu.Visible = True
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
        Exit Sub
    End If

    Select Case State
        Case MENU_STATE_ADDCHAR
            Call SetStatus("Connected, sending character addition data...")

            If newCharSex = SEX_MALE Then
                Call SendAddChar(frmMenu.txtCName.text, SEX_MALE, newCharClass, newCharSprite)
            Else
                Call SendAddChar(frmMenu.txtCName.text, SEX_FEMALE, newCharClass, newCharSprite)
            End If
        Case MENU_STATE_NEWACCOUNT
            Call SetStatus("Connected, sending new account information...")
            Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
        Case MENU_STATE_LOGIN
            Call SetStatus("Connected, sending login information...")
            Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
        Case MENU_STATE_USECHAR
            Call SetStatus("Connected, sending character data...")
            Call SendRequestUseChar
        Case MENU_STATE_NEWCHAR
            Call SetStatus("Connected, requesting new character...")
            Call SendRequestNewChar
        Case MENU_STATE_DELCHAR
            Call SetStatus("Connected, requesting deletion...")
            Call SendRequestDelChar
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
    Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Close form
    frmMain.Hide
    
    ' Set the status
    frmLoad.Show
    Call SetStatus("Leaving the game...")

    ' Reset values
    isLogging = False
    GettingMap = True
    
    ' Destroy TCP
    Set Buffer = New clsBuffer
    Buffer.WriteLong ClientPackets.CQuit
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    Call DestroyTCP
    
    ' destroy temp values
    InGame = False
    MyIndex = 0
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    InvX = 0
    InvY = 0
    EqX = 0
    EqY = 0
    SpellX = 0
    SpellY = 0
    TitleX = 0
    TitleY = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0

    ' unload editors
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    Unload frmEditor_Quest
    Unload frmEditor_Title
    Unload frmMapReport
    Unload frmDebug

    ' hide main form stuffs
    frmMenu.picLogin.Visible = True
    frmMain.txtChat.text = vbNullString
    frmMain.txtMyChat.text = vbNullString
    frmMain.picAdmin.Visible = False
    frmMain.picCurrency.Visible = False
    frmMain.picDialogue.Visible = False
    frmMain.picInventory.Visible = False
    frmMain.picTrade.Visible = False
    frmMain.picBank.Visible = False
    frmMain.picSpells.Visible = False
    frmMain.picCharacter.Visible = False
    frmMain.picOptions.Visible = False
    frmMain.picParty.Visible = False
    frmMain.picQuest.Visible = False
    frmMain.picTitles.Visible = False

    ' Open menu
    frmLoad.Hide
    InMenu = True
    frmMenu.Show

    ' Play menu music
    Stop_Music
    Play_Music Menu_Music
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "logoutGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = 0
    
    ' Set font
    'Call SetFont(FONT_NAME, FONT_SIZE)
    frmMain.Font = "Arial Bold"
    frmMain.FontSize = 10
    
    ' show the main form
    frmLoad.Visible = False
    frmMain.Show
    
    ' Set the focus
    Call SetFocusOnChat
    frmMain.picScreen.Visible = True

    ' get ping
    GetPing

    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.Value = 1
    
    'stop the song playing
    Stop_Music
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Set the status
    frmLoad.Visible = True
    Call SetStatus("Destroying game data...")
        
    ' break out of game loop
    InGame = False
    InMenu = False
    Call DestroyTCP
    
    ' destroy objects in reverse order
    Call DestroyDirectDraw
    Call Destroy_Music

    ' Close forms
    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearGameData()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearPlayers
    Call ClearNpcs
    Call ClearResources
    Call ClearItems
    Call ClearShops
    Call ClearSpells
    Call ClearAnimations
    Call ClearDoors
    Call ClearQuests
    Call ClearTitles

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.text = Txt.text + msg + vbCrLf
    Else
        Txt.text = Txt.text + msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next 'prevent RTE5, no way to handle error

    frmMain.txtMyChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GlobalX As Long
    Dim GlobalY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GlobalX = PB.Left
    GlobalY = PB.top

    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.top = GlobalY + y - SOffsetY
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
            
    ' menu - minimize
    With MenuButton(1)
        .FileName = "minimize"
        .State = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(2)
        .FileName = "exit"
        .State = 0 ' normal
    End With

    ' Render in normaly state
    For i = 1 To MAX_MENUBUTTONS
        renderButton_Menu i
    Next

    ' main - exit
    With MainButton(1)
        .FileName = "minimize"
        .State = 0 ' normal
    End With
    
    ' main - minimize
    With MainButton(2)
        .FileName = "exit"
        .State = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(3)
        .FileName = "inv"
        .State = 0 ' normal
    End With
    
    ' main - skills
    With MainButton(4)
        .FileName = "skills"
        .State = 0 ' normal
    End With
    
    ' main - char
    With MainButton(5)
        .FileName = "char"
        .State = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(6)
        .FileName = "opt"
        .State = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(7)
        .FileName = "trade"
        .State = 0 ' normal
    End With
    
    ' main - party
    With MainButton(8)
        .FileName = "party"
        .State = 0 ' normal
    End With
    
    ' main - quest
    With MainButton(9)
        .FileName = "quest"
        .State = 0 ' normal
    End With
    
    ' main - titles
    With MainButton(10)
        .FileName = "titles"
        .State = 0 ' normal
    End With
    
    ' Render in normaly state
    For i = 1 To MAX_MAINBUTTONS
        RenderButton_Main i
    Next

    ' quest - informations
    With QuestButton(1)
        .FileName = "list"
        .State = 0 ' normal
    End With
    
    ' quest - rewards
    With QuestButton(2)
        .FileName = "rewards"
        .State = 0 ' normal
    End With
    
    ' quest - informations
    With QuestButton(3)
        .FileName = "info"
        .State = 0 ' normal
    End With
    
    ' Render in normaly state
    For i = 1 To MAX_QUESTBUTTONS
        renderButton_Quest i
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).State = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
    Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MenuButton(buttonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).State = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).State = bState
        renderButton_Menu buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).State = 0 'normal
            RenderButton_Main i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Main = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderButton_Main(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MainButton(buttonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).State = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).State = bState
        RenderButton_Main buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' quest specific buttons
Public Sub resetButtons_Quest(Optional ByVal exceptionNum As Long = 0)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_QUESTBUTTONS
        ' only change if different and not exception
        If Not QuestButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            QuestButton(i).State = 0 'normal
            renderButton_Quest i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Quest = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Quest(ByVal buttonNum As Long)
    Dim bSuffix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case QuestButton(buttonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgQuest(buttonNum).Picture = LoadPicture(App.Path & QUESTBUTTON_PATH & QuestButton(buttonNum).FileName & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Quest(ByVal buttonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If QuestButton(buttonNum).State = bState Then Exit Sub
        ' change and render
        QuestButton(buttonNum).State = bState
        renderButton_Quest buttonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Quest", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for debugging
Public Sub DebugAdd(ByVal msg As String, ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Add text
    frmDebug.txtDebug(Index) = frmDebug.txtDebug(Index) & vbNewLine & msg

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DebugAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
