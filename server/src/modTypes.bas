Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec
Public Editor(MAX_EDITORS) As EditorRec
Public TempEditor(1 To MAX_EDITORS) As TempEditorRec

Private Type EditorRec
    Username As String * 20
    Password As String * 20
    
    HasRight(Editor_MaxRights - 1) As Byte
End Type

Private Type TempEditorRec
    InEditor As Byte
    OnIndex As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    buffer As clsBuffer
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    EditorPort As Long
    Website As String
    Scripting As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    Points As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
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
    buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    X As Byte
    Y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    Animation As Long
    Damage As Long
    Level As Long
End Type

Private Type MapNpcRec
    Num As Long
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
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
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    X As Long
    Y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
    Red(0 To 1) As Byte
    Green(0 To 1) As Byte
    Blue(0 To 1) As Byte
    Alpha(0 To 1) As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
    Red(0 To 1) As Byte
    Green(0 To 1) As Byte
    Blue(0 To 1) As Byte
    Alpha(0 To 1) As Byte
End Type
