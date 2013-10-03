Attribute VB_Name = "modEnumerations"
Option Explicit

Public Enum SE_EditorPackets
    SE_AlertMsg = 1
    SE_VersionOK
    ' Make sure SE_MSG_COUNT is below everything else
    SE_MSG_COUNT
End Enum

Public Enum CE_EditorPackets
    CE_LoginUser = 1
    CE_VersionCheck
    CE_SaveDeveloper
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
