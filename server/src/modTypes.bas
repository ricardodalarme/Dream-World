Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(MAX_MAPS) As MapRec
Public ResourceCache(MAX_MAPS) As ResourceCacheRec
Public DoorCache(MAX_MAPS) As DoorCacheRec
Public Player(MAX_PLAYERS) As PlayerRec
Public Bank(MAX_PLAYERS) As BankRec
Public Class() As ClassRec
Public Item(MAX_ITEMS) As ItemRec
Public Npc(MAX_NPCS) As NpcRec
Public MapItem(MAX_MAPS, MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(MAX_MAPS) As MapDataRec
Public Shop(MAX_SHOPS) As ShopRec
Public Spell(MAX_SPELLS) As SpellRec
Public Resource(MAX_RESOURCES) As ResourceRec
Public Animation(MAX_ANIMATIONS) As AnimationRec
Public Party(MAX_PARTYS) As PartyRec
Public Door(MAX_DOORS) As DoorRec
Public Quest(MAX_QUESTS) As QuestRec
Public Title(MAX_TITLES) As TitleRec
Public ChatRoom(MAX_ROOMS) As ChatRoomRec

' server-side stuff
Public MapCache(MAX_MAPS) As Cache
Public PlayersOnMap(MAX_MAPS) As Byte
Public TempPlayer(MAX_PLAYERS) As TempPlayerRec

' options
Public Options As OptionsRec

' Type recs
Public Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    Debug As Byte
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Public Type BankRec
    Item() As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    SType As Byte
End Type

Public Type PlayerQuestRec
    Num As Integer
    Status As Byte
    Part As Byte
End Type

Public Type PlayerTitleRec
    Title() As Long
    Using As Long
End Type

Public Type CharRec
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital() As Long
    ' Stats
    Stat() As Byte
    POINTS As Long
    ' Worn equipment
    Equipment() As Long
    ' Inventory
    Inv() As PlayerInvRec
    Spell() As Long
    ' Hotbar
    Hotbar() As HotbarRec
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' Quest
    Quests() As PlayerQuestRec
    ' NPCs and players killed
    KillNpcs() As Integer
    KillPlayers As Integer
    ' Title
    Title As PlayerTitleRec
End Type

Public Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    ' Characters
    Char() As CharRec
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD() As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer() As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT() As DoTRec
    HoT() As DoTRec
    ' Spell buffer
    SpellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' quest
    QuestInvite As Integer
    QuestSelect As Integer
    ' character
    Char As Byte
    ' chat room
    roomIndex As Long
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
    Npc() As Long
End Type

Public Type ClassRec
    ' Data
    Name As String * NAME_LENGTH
    Stat() As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' Spawn position
    StartMap As Integer
    StartX As Byte
    StartY As Byte
    ' Initial items
    StartItemCount As Long
    StartItem() As Long
    StartValue() As Long
    ' Initial spells
    StarSpellCount As Long
    StarSpell() As Long
    ' Initial titles
    StartTitleCount As Long
    StartTitle() As Long
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
    ' despawn
    canDespawn As Boolean
    despawnTimer As Long
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

Public Type Cache
    Data() As Byte
End Type

Public Type MapNpcRec
    ' Data
    Num As Long
    target As Long
    targetType As Byte
    Vital(Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT() As DoTRec
    HoT() As DoTRec
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

Public Type MapDataRec
    Npc() As MapNpcRec
End Type

Public Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Public Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
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

Public Type AnimationRec
    ' Data
    Name As String * NAME_LENGTH
    Sound As Byte
    Sprite() As Long
    ' timing
    Frames() As Long
    LoopCount() As Long
    LoopTime() As Long
End Type

Public Type MapDoorRec
    ' Localization
    x As Byte
    y As Byte
    ' State
    State As Byte
    ' Respawn time
    RespawnTime As Long
End Type

Public Type DoorCacheRec
    Count As Long
    Data() As MapDoorRec
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
    Message() As String * DESC_LENGTH
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
    color As Byte
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

Public Type PartyRec
    Leader As Long
    Member() As Long
    MemberCount As Long
End Type

Private Type ChatRoomRec
    index As Long
    Name As String
    Members As Long
End Type
