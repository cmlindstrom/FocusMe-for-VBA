VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------
' Anchor Controls. Nov 2006
' Written by Andy Pope © 2006, all rights reserved.
' May be redistributed for free, just keep this comment in place
' BUT may not be sold without the author's explicit permission.
'--------------------------------------------------

Public Cnt As MSForms.control
Public AnchorStyle As enumAnchorStyles
Public MinimumLeft As Single
Public MinimumTop As Single
Public MinimumWidth As Single
Public MinimumHeight As Single
Public OrigLeft As Single
Public OrigTop As Single
Public OrigWidth As Single
Public OrigHeight As Single
