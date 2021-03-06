Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

'  Versioning
Public Const EDITOR_VERSION As String = "0.0.1"

' Filenames
Public Const OPTIONS_FILE As String = "Settings.ini"

' Name Lengths
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
Public Const FONT_NAME As String = "texdefault"
Public Const GFX_EXT As String = ".png"
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32
Public Const TEXTURE_WIDTH = 256

' Render Location nonsense
Public Const TileSetWindowOffSetX As Long = 8
Public Const TileSetWindowOffSetY As Long = 97
Public Const MapViewWindowOffSetX As Long = 286
Public Const MapViewWindowOffSetY As Long = 56

'  Loading states
Public Const LoadStateOptions As String = "Loading Options"
Public Const LoadStateTCP As String = "Initializing TCP"
Public Const LoadStateConnecting As String = "Connecting to Server"
Public Const LoadStateVerCheck As String = "Checking Version"
Public Const LoadStateCheckDir As String = "Checking Directories"
Public Const LoadStateAudio As String = "Initializing Audio"
Public Const LoadStateD3D8 As String = "Initializing D3D8"
Public Const LoadStateLogin As String = "Logging In"
