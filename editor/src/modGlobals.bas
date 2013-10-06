Attribute VB_Name = "modGlobals"
Option Explicit

' Main Loop Global(s)
Public EditorLooping As Boolean

'  Loading Bar Globals
Public LoadBarPerc As Byte
Public LoadBarWidth As Long

'  Loading String Globals
Public Const TestText As String = "Initializing Engine"

' TCP Globals
Public EditorBuffer As clsBuffer

' Misc Stuff
Public CheckingVersion As Boolean
Public LoggingIn As Boolean

' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte

'  Render Location Globals
Public RecTileSelectWindow As D3DRECT
Public TileSelectHeight As Long
Public MapViewWindow As D3DRECT
Public MapViewHeight As Long
Public MapViewWidth As Long
Public MapViewTileOffSetX As Long
Public MapViewTileOffSetY As Long

' Map Editor GLobals
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorTileX As Long
Public EditorTileY As Long
Public MouseX As Long
Public MouseY As Long
Public MouseXMove As Long
Public MouseYMove As Long
Public CurrentMap As Long
Public HasMapChanged As Boolean
Public ShowMouse As Boolean
Public CurX As Long
Public CurY As Long

' MAX Globals
Public MAX_MAPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_ANIMATIONS As Long
Public MAX_INV As Long
Public MAX_MAP_ITEMS As Long
Public MAX_MAP_NPCS As Long
Public MAX_SHOPS As Long
Public MAX_PLAYER_SPELLS As Long
Public MAX_SPELLS As Long
Public MAX_RESOURCES As Long
Public MAX_LEVELS As Long
Public MAX_BANK As Long
Public MAX_HOTBAR As Long

