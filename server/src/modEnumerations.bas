Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CSearch
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public Enum SE_EditorPackets
    SE_AlertMsg = 1
    SE_VersionOK
    SE_LoginOK
    SE_MapNames
    SE_MapData
    SE_ResourceData
    SE_MaxAmounts
    ' Make sure SE_MSG_COUNT is below everything else
    SE_MSG_COUNT
End Enum

Public Enum CE_EditorPackets
    CE_LoginUser = 1
    CE_VersionCheck
    CE_SaveDeveloper
    CE_RequestMap
    ' Make sure CE_MSG_COUNT is below everything else
    CE_MSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long
Public HandleEditorDataSub(CE_MSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
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

Public Enum MapMorals
    MapMoralNone = 0
    MapMoralSafe
End Enum

Public Enum ItemTypes
    ItemTypeNone = 0
    ItemTypeWeapon
    ItemTypeArmor
    ItemTypeHelmet
    ItemTypeShield
    ItemTypeConsume
    ItemTypeKey
    ItemTypeCurrency
    ItemTypeSpell
    ItemTypeScripted
End Enum

Public Enum Genders
    SexMale = 0
    SexFemale
End Enum

Public Enum Directions
    North = 0
    South
    West
    East
End Enum

Public Enum PlayerRanks
    RankPlayer = 0
    RankModerator
    RankMapper
    RankDeveloper
    RankAdministrator
End Enum

Public Enum NPCTypes
    NPCTypeAggressive = 0
    NPCTypeNeutral
    NPCTypeFriendly
    NPCTypeStationary
    NPCTypeProtectAllies
    NPCTypeScripted
End Enum

Public Enum SpellTypes
    SpellTypeDamageHP = 0
    SpellTypeDamageMP
    SpellTypeHealHP
    SpellTypeHealMP
    SpellTypeWarp
    SpellTypeScripted
End Enum

Public Enum ActionMessages
    ActionMsgStatic = 0
    ActionMsgScroll
    ActionMsgScreen
End Enum

Public Enum TargetTypes
    TargetTypeNone = 0
    TargetTypePlayer
    TargetTypeNPC
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
    BrightGreen
    BrightCyan
    BrightRed
    Pink
    Yellow
    White
    DarkBrown
    Orange
End Enum

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

Public Enum EditorTypes
    EditorMap = 1
    EditorNPC
    EditorPlayer
    EditorResource
    EditorSpell
    EditorItem
    EditorShop
    EditorAnimation
    EditorDeveloper
End Enum
