Attribute VB_Name = "modDirectX8"
Option Explicit

' ********************
' Credit goes to Sekaru for this Module.
' www.sleepystudios.com
' Original file(s) can be found at: http://www.eclipseorigins.com/community/index.php?/topic/130272-clean-optimised-dx8/
' ********************

'*********************
'* DX8 Functionality *
'*********************

' Main DirectX root
Public DX As DirectX8
Public Direct3D8 As Direct3D8
Public D3DX8 As D3DX8

' Main DirectX visuals
Public D3DDevice8 As Direct3DDevice8
Public DisplayMode As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS

Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    RHW As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

' Texture system
Public D3DT_TEXTURE() As TextureRec
Private Const TEXTURE_NULL As Long = 0

Private Type TextureRec
    Texture As Direct3DTexture8
    Width As Long
    Height As Long
    Path As String
    UnloadTimer As Long
    Loaded As Boolean
End Type

' Textures
Public Tex_TileSet() As Long
Public Tex_Character() As Long
Public Tex_Paperdoll() As Long
Public Tex_Item() As Long
Public Tex_Resource() As Long
Public Tex_Animation() As Long
Public Tex_SpellIcon() As Long
Public Tex_Face() As Long
Public Tex_Blood As Long
Public Tex_DirBlock As Long
Public Tex_Outline As Long

' Texture counts
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long

' Texture values
Private TextureNum As Long
Public CurrentTexture As Long

' Used for setting textures
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long

    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Public Sub InitDirect3D8()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Create the DirectX
    Set DX = New DirectX8
    Set Direct3D8 = DX.Direct3DCreate()
    Set D3DX8 = New D3DX8
    
    ' Find the best processing speed
    If Not InitD3DDevice8() Then
        Call MsgBox("DirectX8 had trouble initiating. Please make sure your graphics card can support DirectX8 and/or is installed.")
        Call UnloadDirectX
        End
    End If
    
    ' Cache the textures
    Call CacheTextures
    
    ' Begin initialising the full engine
    Call InitRenderStates
    Call EngineInitFontTextures
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "InitDirect3D8", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitRenderStates()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With D3DDevice8
        ' Set the shader to be used
        .SetVertexShader FVF
    
        ' Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        ' Particle engine settings
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        ' Set the texture stage stats (filters)
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With

' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "InitRenderStates", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function InitD3DDevice8() As Boolean
    On Error GoTo loadError
    
    Call Direct3D8.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DisplayMode)
    
    D3DWindow.Windowed = True
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DisplayMode.Format
    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    
    Select Case Options.Device
        ' Hardware Rendering
        Case 1
            Set D3DDevice8 = Direct3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
            InitD3DDevice8 = True
            Exit Function
        ' Software Rendering
        Case 2
            Set D3DDevice8 = Direct3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
            InitD3DDevice8 = True
            Exit Function
        ' Mixed Rendering
        Case 3
            Set D3DDevice8 = Direct3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
            InitD3DDevice8 = True
            Exit Function
        ' Pure Rendering
        Case 4
            Set D3DDevice8 = Direct3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_PUREDEVICE, D3DWindow)
            InitD3DDevice8 = True
            Exit Function
    End Select

loadError:
    Set D3DDevice8 = Nothing
    InitD3DDevice8 = False
End Function

Public Sub UnloadDirectX()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear the objects
    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    If Not Direct3D8 Is Nothing Then Set Direct3D8 = Nothing
    
    ' Clear the textures
    For i = 0 To TextureNum
        Set D3DT_TEXTURE(i).Texture = Nothing
    Next
    
    ' Clear the master object
    If Not DX Is Nothing Then Set DX = Nothing

' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "UnloadDirectX", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function SetTexturePath(ByVal Path As String) As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TextureNum = TextureNum + 1
    ReDim Preserve D3DT_TEXTURE(0 To TextureNum) As TextureRec
    
    D3DT_TEXTURE(TextureNum).Path = Path
    SetTexturePath = TextureNum
    D3DT_TEXTURE(TextureNum).Loaded = False
    
' Do not put any code beyond this line, this is the error handler.
    Exit Function
errorhandler:
    HandleError "SetTexturePath", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTexture(ByVal TextureNum As Long)
Dim Tex_Info As D3DXIMAGE_INFO_A
Dim Path As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Find the texture's path
    Path = D3DT_TEXTURE(TextureNum).Path
    
    Select Case D3DT_TEXTURE(TextureNum).Width
        Case 0
            Set D3DT_TEXTURE(TextureNum).Texture = D3DX8.CreateTextureFromFileEx(D3DDevice8, Path, _
            D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_NONE, -65281, Tex_Info, ByVal 0)
            
            ' Get the size of the D3DXIMAGE RECT
            D3DT_TEXTURE(TextureNum).Height = Tex_Info.Height
            D3DT_TEXTURE(TextureNum).Width = Tex_Info.Width
        Case Is > 0
            Set D3DT_TEXTURE(TextureNum).Texture = D3DX8.CreateTextureFromFileEx(D3DDevice8, Path, _
            D3DT_TEXTURE(TextureNum).Width, D3DT_TEXTURE(TextureNum).Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
            D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, -65281, ByVal 0, ByVal 0)
    End Select
    
    ' Re-set the texture unloading timer
    D3DT_TEXTURE(TextureNum).UnloadTimer = GetTickCount
    D3DT_TEXTURE(TextureNum).Loaded = True
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures()
Dim Count As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Count = UBound(D3DT_TEXTURE)
    If Count <= 0 Then Exit Sub
    
    For i = 0 To Count
        With D3DT_TEXTURE(i)
            If .Loaded = True Then ' <--- Missing from the actual base. Seemed too important to leave out.
                If .UnloadTimer > GetTickCount + 60000 Then
            
                    ' Clear it from the memory
                    Set .Texture = Nothing
                    Call ZeroMemory(ByVal VarPtr(D3DT_TEXTURE(i)), LenB(D3DT_TEXTURE(i)))
                
                    ' Set it to unloaded
                .UnloadTimer = 0
                    .Loaded = False
                End If
            End If
        End With
    Next
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "UnloadTextures", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetTexture(ByVal Texture As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Texture <> CurrentTexture Then
        ' Find the texture and make sure ti exists
        If Texture > UBound(D3DT_TEXTURE) Then Texture = UBound(D3DT_TEXTURE)
        If Texture < 0 Then Texture = 0
        
        ' Check if the texture is loaded
        If Not Texture = TEXTURE_NULL Then
            ' Reload it from the path
            If Not D3DT_TEXTURE(Texture).Loaded Then Call LoadTexture(Texture)
        End If
        
        ' Set the current texture
        Call D3DDevice8.SetTexture(0, D3DT_TEXTURE(Texture).Texture)
        CurrentTexture = Texture
    End If
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "SetTexture", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheTextures()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Cache the Tilesets
    NumTileSets = 1
    Do While FileExist(App.Path & GFX_PATH & "Tilesets\" & NumTileSets & GFX_EXT, True)
        ReDim Preserve Tex_TileSet(0 To NumTileSets)
        Tex_TileSet(NumTileSets) = SetTexturePath(App.Path & GFX_PATH & "Tilesets\" & NumTileSets & GFX_EXT)
        NumTileSets = NumTileSets + 1
    Loop
    NumTileSets = NumTileSets - 1

    ' Cache the Characters
    NumCharacters = 1
    Do While FileExist(App.Path & GFX_PATH & "characters\" & NumCharacters & GFX_EXT, True)
        ReDim Preserve Tex_Character(0 To NumCharacters)
        Tex_Character(NumCharacters) = SetTexturePath(App.Path & GFX_PATH & "characters\" & NumCharacters & GFX_EXT)
        NumCharacters = NumCharacters + 1
    Loop
    NumCharacters = NumCharacters - 1
    
    ' Cache the Paperdolls
    NumPaperdolls = 1
    Do While FileExist(App.Path & GFX_PATH & "Paperdolls\" & NumPaperdolls & GFX_EXT, True)
        ReDim Preserve Tex_Paperdoll(0 To NumPaperdolls)
        Tex_Paperdoll(NumPaperdolls) = SetTexturePath(App.Path & GFX_PATH & "Paperdolls\" & NumPaperdolls & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
    Loop
    NumPaperdolls = NumPaperdolls - 1

    ' Cache the Items
    NumItems = 1
    Do While FileExist(App.Path & GFX_PATH & "Items\" & NumItems & GFX_EXT, True)
        ReDim Preserve Tex_Item(0 To NumItems)
        Tex_Item(NumItems) = SetTexturePath(App.Path & GFX_PATH & "Items\" & NumItems & GFX_EXT)
        NumItems = NumItems + 1
    Loop
    NumItems = NumItems - 1
    
    ' Cache the Resources
    NumResources = 1
    Do While FileExist(App.Path & GFX_PATH & "Resources\" & NumResources & GFX_EXT, True)
        ReDim Preserve Tex_Resource(0 To NumResources)
        Tex_Resource(NumResources) = SetTexturePath(App.Path & GFX_PATH & "Resources\" & NumResources & GFX_EXT)
        NumResources = NumResources + 1
    Loop
    NumResources = NumResources - 1

    ' Cache the Animations
    NumAnimations = 1
    Do While FileExist(App.Path & GFX_PATH & "animations\" & NumAnimations & GFX_EXT, True)
        ReDim Preserve Tex_Animation(0 To NumAnimations)
        Tex_Animation(NumAnimations) = SetTexturePath(App.Path & GFX_PATH & "animations\" & NumAnimations & GFX_EXT)
        NumAnimations = NumAnimations + 1
    Loop
    NumAnimations = NumAnimations - 1
    
    ' Cache the SpellIcons
    NumSpellIcons = 1
    Do While FileExist(App.Path & GFX_PATH & "SpellIcons\" & NumSpellIcons & GFX_EXT, True)
        ReDim Preserve Tex_SpellIcon(0 To NumSpellIcons)
        Tex_SpellIcon(NumSpellIcons) = SetTexturePath(App.Path & GFX_PATH & "SpellIcons\" & NumSpellIcons & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
    Loop
    NumSpellIcons = NumSpellIcons - 1
    
    ' Cache the Faces
    NumFaces = 1
    Do While FileExist(App.Path & GFX_PATH & "Faces\" & NumFaces & GFX_EXT, True)
        ReDim Preserve Tex_Face(0 To NumFaces)
        Tex_Face(NumFaces) = SetTexturePath(App.Path & GFX_PATH & "Faces\" & NumFaces & GFX_EXT)
        NumFaces = NumFaces + 1
    Loop
    NumFaces = NumFaces - 1
    
    ' Now this is where we'll start caching and pre-loading some of the "required" textures, such as the blood and target ones.
    ' Nothing too complicated, just making sure they exist and are loaded without any fuss.
    Tex_Blood = SetTexturePath(App.Path & GFX_PATH & "blood" & GFX_EXT)
    Call LoadTexture(Tex_Blood)
    ' A little special touch for the blood, we need to know how many blood textures we have! So let's calculate it.
    BloodCount = D3DT_TEXTURE(Tex_Blood).Width / PIC_X
    
    Tex_DirBlock = SetTexturePath(App.Path & GFX_PATH & "direction" & GFX_EXT)
    Call LoadTexture(Tex_DirBlock)
    
    Tex_Outline = SetTexturePath(App.Path & GFX_PATH & "outline" & GFX_EXT)
    Call LoadTexture(Tex_Outline)
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "CacheTextures", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderGraphic(ByVal Texture As Long, x As Long, y As Long, DW As Long, DH As Long, Optional TW As Long, Optional TH As Long, _
Optional OX As Long, Optional OY As Long, Optional R As Byte = 255, Optional G As Byte = 255, Optional B As Byte = 255, Optional A As Byte = 255)
Dim Box(0 To 3) As TLVERTEX, i As Long, TextureWidth As Long, TextureHeight As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' set the texture
    Call SetTexture(Texture)
    
    If TW = 0 Then TW = DW
    If TH = 0 Then TH = DH
    
    ' set the texture size
    TextureWidth = D3DT_TEXTURE(Texture).Width
    TextureHeight = D3DT_TEXTURE(Texture).Height
    
    ' exit out if we need to
    If Texture <= 0 Or TextureWidth <= 0 Or TextureHeight <= 0 Then Exit Sub
    
    For i = 0 To 3
        Box(i).RHW = 1
        Box(i).color = D3DColorRGBA(R, G, B, A)
    Next

    Box(0).x = x
    Box(0).y = y
    Box(0).tu = (OX / TextureWidth)
    Box(0).tv = (OY / TextureHeight)
    Box(1).x = x + DW
    Box(1).tu = (OX + TW + 1) / TextureWidth
    Box(2).x = Box(0).x
    Box(3).x = Box(1).x

    Box(2).y = y + DH
    Box(2).tv = (OY + TH + 1) / TextureHeight

    Box(1).y = Box(0).y
    Box(1).tv = Box(0).tv
    Box(2).tu = Box(0).tu
    Box(3).y = Box(2).y
    Box(3).tu = Box(1).tu
    Box(3).tv = Box(2).tv
    
    Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), FVF_Size)
    D3DT_TEXTURE(Texture).UnloadTimer = GetTickCount
    
' Do not put any code beyond this line, this is the error handler.
    Exit Sub
errorhandler:
    HandleError "RenderGraphic", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

