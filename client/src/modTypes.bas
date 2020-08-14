Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public Player(MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(MAX_ITEMS) As ItemRec
Public Npc(MAX_NPCS) As NpcRec
Public MapItem(MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(MAX_MAP_NPCS) As MapNpcRec
Public Shop(MAX_SHOPS) As ShopRec
Public Spell(MAX_SPELLS) As SpellRec
Public Resource(MAX_RESOURCES) As ResourceRec
Public Animation(MAX_ANIMATIONS) As AnimationRec
Public Door(MAX_DOORS) As DoorRec
Public Quest(MAX_QUESTS) As QuestRec
Public Title(MAX_TITLES) As TitleRec

' client-side stuff
Public ActionMsg(MAX_BYTE) As ActionMsgRec
Public Blood(MAX_BYTE) As BloodRec
Public AnimInstance(MAX_BYTE) As AnimInstanceRec
Public MenuButton(MAX_MENUBUTTONS) As ButtonRec
Public MainButton(MAX_MAINBUTTONS) As ButtonRec
Public QuestButton(MAX_QUESTBUTTONS) As ButtonRec
Public CharData(MAX_PLAYER_CHARS) As CharDataRec
Public Party As PartyRec

' options
Public Options As OptionsRec

' Type recs
Public Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    Music As Byte
    Sound As Byte
    Debug As Byte
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Public Type BankRec
    Item(MAX_BANK) As PlayerInvRec
End Type

Public Type PlayerQuestRec
    Num As Integer
    Status As Byte
    Part As Byte
End Type

Public Type PlayerTitleRec
    Title(MAX_PLAYER_TITLES) As Long
    Using As Long
End Type

Public Type PlayerRec
    ' General
    Name As String
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital() As Long
    MaxVital() As Long
    ' Stats
    Stat() As Byte
    POINTS As Long
    ' Worn equipment
    Equipment() As Long
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' Quests
    Quests() As PlayerQuestRec
    ' Npc and players killed
    KillNpcs() As Integer
    KillPlayers As Integer
    ' Title
    Title As PlayerTitleRec
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
End Type

Public Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(MapLayer.Layer_Count - 1, MAX_MAP_LAYERS) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Public Type MapRec
    ' Data
    Name As String * NAME_LENGTH
    Music As Byte
    Revision As Long
    Moral As Byte
    ' Panorama
    Panorama As Long
    ' Tint
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    ' Fog
    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte
    ' Links
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    UpLeft As Long
    UpRight As Long
    DownLeft As Long
    DownRight As Long
    ' Boot
    BootMap As Long
    BootX As Byte
    BootY As Byte
    ' Maz sizes
    MaxX As Byte
    MaxY As Byte
    ' Tiles
    Tile() As TileRec
    Npc(MAX_MAP_NPCS) As Long
End Type

Public Type ClassRec
    Name As String * NAME_LENGTH
    Stat(Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
End Type

Public Type ItemRec
    ' Data
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    Sound As Byte
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ' Requirements
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Stat_Req() As Byte
    ' Specifications
    Price As Long
    Rarity As Byte
    ' Equipments data
    Animation As Long
    Paperdoll As Long
    Damage As Long
    Protection As Long
    Add_Stat() As Byte
    Speed As Long
    ' Counsume data
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    ' Surprise
    BagItem() As Long
    BagValue() As Long
End Type

Public Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    y As Byte
End Type

Public Type NpcRec
    ' Data
    Name As String * NAME_LENGTH
    AttackSay As String * SAY_LENGTH
    Sound As Byte
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    ' Drop
    DropChance() As Byte
    DropItem() As Integer
    DropItemValue() As Long
    ' Data used in a fight
    Stat() As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    ' Others
    Quest() As Integer
    ShopNum As Long
End Type

Public Type MapNpcRec
    ' Data
    Num As Long
    target As Long
    TargetType As Byte
    Vital(Vitals.Vital_Count - 1) As Long
    ' Localization
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
End Type

Public Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Public Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem() As TradeItemRec
End Type

Public Type SpellRec
    ' Data
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    Icon As Long
    Sound As Byte
    Type As Byte
    ' Effects when using
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    CastTime As Long
    CDTime As Long
    MPCost As Long
    ' Warp
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    ' Effects
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    BaseStat As Byte
    ' Requirements
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
End Type

Public Type MapResourceRec
    x As Long
    y As Long
    ResourceState As Byte
End Type

Public Type ResourceRec
    ' Data
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As Byte
    Type As Byte
    Animation As Long
    Health As Long
    RespawnTime As Long
    ' Images
    ResourceImage As Long
    ExhaustedImage As Long
    ' Rewards
    ItemReward As Long
    ' Requirements
    ToolRequired As Long
End Type

Public Type ActionMsgRec
    message As String
    Created As Long
    Type As Long
    Color As Long
    Scroll As Long
    x As Long
    y As Long
    Timer As Long
End Type

Public Type BloodRec
    Sprite As Long
    Timer As Long
    x As Long
    y As Long
End Type

Public Type AnimationRec
    ' Data
    Name As String * NAME_LENGTH
    Sound As Byte
    Sprite() As Long
    ' timing
    Frames() As Long
    LoopCount() As Long
    looptime() As Long
End Type

Public Type AnimInstanceRec
    ' Data
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(1) As Long
    ' rendering check
    Used(1) As Boolean
    ' counting the loop
    LoopIndex(1) As Long
    FrameIndex(1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    FileName As String
    State As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type MapDoorRec
    ' Localization
    x As Byte
    y As Byte
    ' State
    State As Byte
End Type

Public Type DoorRec
    ' Data
    Name As String * NAME_LENGTH
    OpeningImage As Integer
    ClosedImage As Integer
    OpenWith As Integer
    Respawn As Long
    Animation As Long
    Sound As Byte
    ' Warp
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    ' Requirements
    LevelReq As Integer
    Stat_Req() As Byte
End Type

Public Type QuestTaskRec
    ' Data
    Type As Byte
    message() As String * SAY_LENGTH
    Instant As Boolean
    ' Objectives
    Num As Integer
    Value As Long
End Type

Public Type QuestRec
    ' Data
    Name As String * NAME_LENGTH
    Description As String * DESC_LENGTH
    Retry As Boolean
    ' Requirements
    LevelReq As Integer
    StatReq() As Long
    QuestReq As Integer
    ClassReq As Byte
    SpriteReq As Integer
    ' Rewards
    LevelRew As Integer
    ExpRew As Long
    StatRew() As Long
    VitalRew() As Long
    ClassRew As Byte
    SpriteRew As Integer
    ' Tasks
    Task() As QuestTaskRec
End Type

Public Type TitleRec
    ' Data
    Name As String * NAME_LENGTH
    Description As String * DESC_LENGTH
    Icon As Integer
    Type As Byte
    Color As Byte
    Sound As Byte
    UseAnimation As Long
    RemoveAnimation As Long
    Passive As Boolean
    ' Requirements
    LevelReq As Long
    StatReq() As Long
    ' Rewards
    StatRew() As Long
    VitalRew() As Long
End Type

Public Type CharDataRec
    ' Data
    Name As String * NAME_LENGTH
    Level As Long
    Class As Long
End Type

