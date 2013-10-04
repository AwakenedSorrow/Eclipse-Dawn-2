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

' MAX Globals
Public MAX_MAPS As Long

