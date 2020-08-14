Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 3
Public Const CLIENT_REVISION As Byte = 0

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' ********************************************************
' * The values below must match with the client' s values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Long = 50
Public Const MAX_ITEMS As Long = 75
Public Const MAX_NPCS As Long = 75
Public Const MAX_ANIMATIONS As Long = 75
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 200
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 75
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 75
Public Const MAX_LEVELS As Long = 100
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_MAP_LAYERS As Long = 3
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_DOORS As Byte = 75
Public Const MAX_QUESTS As Byte = 75
Public Const MAX_PLAYER_QUESTS As Byte = 75
Public Const MAX_QUEST_TASKS As Byte = 5
Public Const MAX_NPC_QUESTS As Byte = 5
Public Const MAX_TITLES As Integer = 100
Public Const MAX_PLAYER_TITLES As Byte = 30
Public Const MAX_BAG As Long = 15
Public Const MAX_PLAYER_CHARS As Byte = 3
Public Const MAX_ROOMS As Long = 20

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 1:30 seconds
Public Const MAX_DOTS As Long = 30

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12
Public Const DESC_LENGTH As Byte = 200
Public Const SAY_LENGTH As Byte = 100

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 50
Public Const MAX_MAPX As Byte = 23
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEYOPEN As Byte = 5
Public Const TILE_TYPE_RESOURCE As Byte = 6
Public Const TILE_TYPE_DOOR As Byte = 7
Public Const TILE_TYPE_NPCSPAWN As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_BANK As Byte = 10
Public Const TILE_TYPE_HEAL As Byte = 11
Public Const TILE_TYPE_TRAP As Byte = 12
Public Const TILE_TYPE_SLIDE As Byte = 13

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP_LEFT As Byte = 4
Public Const DIR_UP_RIGHT As Byte = 5
Public Const DIR_DOWN_LEFT As Byte = 6
Public Const DIR_DOWN_RIGHT As Byte = 7

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3
Public Const DIALOGUE_TYPE_QUEST As Byte = 4

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' Quest constants
Public Const QUEST_TYPE_NONE As Byte = 0
Public Const QUEST_TYPE_KILLNPC As Byte = 1
Public Const QUEST_TYPE_KILLPLAYER As Byte = 2
Public Const QUEST_TYPE_GOTOMAP As Byte = 3
Public Const QUEST_TYPE_TALKNPC As Byte = 4
Public Const QUEST_TYPE_COLLECTITEMS As Byte = 5

' Quest status constants
Public Const QUEST_STATUS_NONE As Byte = 0
Public Const QUEST_STATUS_STARTING As Byte = 1
Public Const QUEST_STATUS_COMPLETE As Byte = 2
Public Const QUEST_STATUS_END As Byte = 3

' Titles const
Public Const TITLE_TYPE_NORMAL As Byte = 0
Public Const TITLE_TYPE_INITIAL As Byte = 1
