Attribute VB_Name = "modGlobals"
Option Explicit

' Paperdoll rendering order
Public PaperdollOrder() As Long

' music & sound list cache
Public musicCache() As String
Public soundCache() As String
Public hasPopulated As Boolean

' global dialogue index
Public dialogueIndex As Long
Public dialogueData1 As Long

' Buttons
Public LastButtonSound_Menu As Long
Public LastButtonSound_Main As Long
Public LastButtonSound_Quest As Long

' Hotbar
Public Hotbar(MAX_HOTBAR) As HotbarRec

' Amount of blood decals
Public BloodCount As Long

' main menu unloading
Public EnteringGame As Boolean

' GUI
Public HPBar_Width As Long
Public SPRBar_Width As Long
Public EXPBar_Width As Long

' Party GUI
Public Party_HPWidth As Long
Public Party_SPRWidth As Long

' targetting
Public myTarget As Long
Public myTargetType As Long

' for directional blocking
Public DirArrowX(4) As Byte
Public DirArrowY(4) As Byte

' trading
Public TradeTimer As Long
Public InTrade As Long
Public TradeYourOffer(MAX_INV) As PlayerInvRec
Public TradeTheirOffer(MAX_INV) As PlayerInvRec
Public TradeX As Long
Public TradeY As Long

' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean

' Cache the doors in an array
Public MapDoor() As MapDoorRec
Public Door_Index As Long
Public Doors_Init As Boolean

' map editor boxes
Public shpSelectedTop As Long
Public shpSelectedLeft As Long
Public shpSelectedHeight As Long
Public shpSelectedWidth As Long
Public shpLocTop As Long
Public shpLocLeft As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long
Public EditorDoor As Long
Public EditorEvent As Long
Public QuestTask As Byte
Public EventsList As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map Resources
Public ResourceEditorNum As Long

' Used for map editor heal & trap & slide tiles
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long

' inv drag + drop
Public DragInvSlotNum As Long
Public InvX As Long
Public InvY As Long

' bank drag + drop
Public DragBankSlotNum As Long
Public BankX As Long
Public BankY As Long

' spell drag + drop
Public DragSpell As Long

' title globals
Public DragTitle As Long
Public TitleX As Long
Public TitleY As Long
Public LastTitleDesc As Long ' Stores the last title we showed in desc

' gui
Public EqX As Long
Public EqY As Long
Public SpellX As Long
Public SpellY As Long
Public InvItemFrame(MAX_INV) As Byte  ' Used for animated items
Public LastItemDesc As Long ' Stores the last item we showed in desc
Public LastSpellDesc As Long ' Stores the last spell we showed in desc
Public LastBankDesc As Long ' Stores the last bank item we showed in desc
Public tmpCurrencyItem As Long
Public InShop As Long ' is the player in a shop?
Public ShopAction As Byte ' stores the current shop action
Public InBank As Long
Public CurrencyMenu As Byte

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(MAX_INV) As PlayerInvRec    ' Inventory
Public PlayerSpells(MAX_PLAYER_SPELLS) As Long
Public InventoryItemSelected As Long
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Game text buffer
Public MyText As String

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean
Public InMenu As Boolean

' Text variables
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public DirUpLeft As Boolean
Public DirUpRight As Boolean
Public DirDownLeft As Boolean
Public DirDownRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Used for dragging Picture Boxes
Public SOffsetX As Long
Public SOffsetY As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long

' Mouse cursor tile location
Public CurX As Long
Public CurY As Long

' Game editors
Public CurrentEditor As Byte
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

' Maximum classes
Public Max_Classes As Long
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' Tick count interval
Public TickInterval As Long

' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte

' fps lock
Public FPS_Lock As Boolean

' New char
Public newCharSprite As Long
Public newCharClass As Long
Public newCharSex As Byte

' Editor edited items array
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public NPC_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_RESOURCES) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Door_Changed(1 To MAX_DOORS) As Boolean
Public Quest_Changed(1 To MAX_QUESTS) As Boolean
Public Title_Changed(1 To MAX_TITLES) As Boolean

' looping saves
Public Player_HighIndex As Long
Public Npc_HighIndex As Long
Public Action_HighIndex As Long

' Tips
Public Max_Tips As Byte
Public Tip() As String * DESC_LENGTH

' Use in screen shot map
Public ScreenShot As Boolean

' fog
Public fogOffsetX As Long
Public fogOffsetY As Long
