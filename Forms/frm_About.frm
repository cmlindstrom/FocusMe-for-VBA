VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_About 
   Caption         =   "About"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   OleObjectBlob   =   "frm_About.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "frm_About"

' Event Handlers


Private Sub UserForm_Activate()
    ' Fires every time the Window gets user focus
    ' SetFormPosition Me, 100, 100
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Closing
End Sub

Private Sub btn_OK_Click()
    Unload Me
End Sub

' Constructor

Private Sub UserForm_Initialize()

    Me.Caption = Commands.AppName & " - About"
    
    Me.lbl_Version = "Version " & Commands.AppVersion
    
    Me.txtbx_Description.text = GetDescription

End Sub

Private Sub UserForm_Terminate()
    ' Clean Up
End Sub

' Methods

' Supporting Methods

Private Function GetDescription() As String

    Dim strText As String
    
    strText = "FocusMe for Outlook (VBA) - 2018-2020" & vbCrLf & vbCrLf
    strText = strText & "FocusMe helps people focus their everyday activities " & _
                "to accomplish what is most important."
                
    strText = strText & vbCrLf & vbCrLf
                
    strText = strText & "For technical questions and customer support, " & _
                "send an email to support@ceptara.com"
    
    GetDescription = strText

End Function



