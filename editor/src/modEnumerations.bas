Attribute VB_Name = "modEnumerations"
Option Explicit

Public Enum SE_EditorPackets
    SE_AlertMsg = 1
    SE_VersionOK
    SE_LoginOK
    SE_MapNames
    SE_MapData
    SE_ResourceData
    SE_MaxAmounts
    SE_AnimationData
    SE_SpellData
    SE_ShopData
    SE_MapSaved
    ' Make sure SE_MSG_COUNT is below everything else
    SE_MSG_COUNT
End Enum

Public Enum CE_EditorPackets
    CE_LoginUser = 1
    CE_VersionCheck
    CE_SaveDeveloper
    CE_RequestMap
    CE_SaveMap
    ' Make sure CE_MSG_COUNT is below everything else
    CE_MSG_COUNT
End Enum

Public HandleDataSub(SE_MSG_COUNT) As Long

Public Enum EditorRights
    CanEditMaps = 1
    CanUseChat
    CanOpenDatabase
    CanChangeOwnDetails
    CanEditNPC
    CanEditMap
    CanEditShop
    CanEditItem
    CanEditSpell
    CanEditResource
    CanEditPlayer
    CanEditAnimation
    CanEditScript
    CanEditDeveloper
    CanAddDeveloper
    CanRemoveDeveloper
    ' Always below everything else!
    Editor_MaxRights
End Enum

Public Enum MapLayer
    Ground = 1
    mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

Public Enum Colors
    Black = 0
    Blue
    Green
    Cyan
    Red
    Magenta
    Brown
    Grey
    DarkGrey
    BrightBlue
    brightgreen
    BrightCyan
    BrightRed
    Pink
    yellow
    White
    DarkBrown
    Orange
End Enum

Public Enum TileTypes
    TileTypeWalkable = 0
    TileTypeBlocked
    TileTypeWarp
    TileTypeItem
    TileTypeNPCAvoid
    TileTypeKey
    TileTypeKeyOpen
    TileTypeResource
    TileTypeDoor
    TileTypeNPCSpawn
    TileTypeShop
    TileTypeBank
    TileTypeHeal
    TileTypeTrap
    TileTypeSlide
    TileTypeScripted
End Enum
