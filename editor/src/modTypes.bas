Attribute VB_Name = "modTypes"
Option Explicit

Public Options As OptionsRec
Public Editor As EditorRec
Public TempEditor As EditorRec
Public Map As MapRec

Public Resource() As ResourceRec
Public Animation() As AnimationRec
Public Shop() As ShopRec
Public Spell() As SpellRec

'  **************************************
'  **************************************
'  **************************************

Type OptionsRec
    '  Server Related
    ServerIP As String
    ServerPort As Long
    
    ' Account
    RememberUser As Byte
    Username As String
    
    ' Tileset Options
    TileSetName() As String
    
    ' Debug
    device As Byte
End Type

Private Type EditorRec
    Username As String * NAME_LENGTH
    
    HasRight(Editor_MaxRights - 1) As Byte
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type TileRec
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
    Npc() As Long
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

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type


Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem() As TradeItemRec
End Type
