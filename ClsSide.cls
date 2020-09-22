VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum P_State
    vbOpen = 0
    vbClose = 1
End Enum

Enum P_Type
    vbMain = 0
    vbSub = 1
End Enum

Enum P_Style
    xpHyperlink = 0
    xpCustom = 1
End Enum

Type POINTAPI
    x As Long
    y As Long
End Type

Enum StateConstants
    btDown = 0
    btUp = 1
    btOver = 2
End Enum
