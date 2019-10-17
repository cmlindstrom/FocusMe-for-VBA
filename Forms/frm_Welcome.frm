VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Welcome 
   Caption         =   "Welcome"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frm_Welcome.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "frm_Welcome"

Dim st As Setup
Attribute st.VB_VarHelpID = -1

' - Events

' - Properties

Dim f_step As Integer

' - Event Handlers

Private Sub btn_Execute_Click()

    Dim i As Integer

    If f_step = 3 Then
        Unload Me
    End If
    If f_step = 2 Then
        Status "Loading projects from Outlook categories..."
        i = st.ImportProjectsFromCategories
        
        If i >= 0 Then
            AppendText " - - imported " & i & " projects from the master category list."
            FinalStep
        Else
            AppendText " - - encountered an error while importing."
            ErrorStep
        End If

        Status
        
    End If
    If f_step = 1 Then
        Status "Indexing Outlook folders..."
        i = st.IndexOutlookFolders
        
        If i > 0 Then
            AppendText " - - found " & i & " folders."
            StepTwo
        Else
            AppendText " - - encountered an error while indexing."
            ErrorStep
        End If
        
        Status

    End If


End Sub

Private Sub btn_Cancel_Click()
    Unload Me

    'If f_step = 2 Then
    '    ' Exit dialog
    '    Unload Me
    'End If

End Sub

' - Constructor

Private Sub UserForm_Initialize()
    Me.Caption = Commands.AppName & " - Welcome"
    
    ' For now
    Me.btn_About.Visible = False
    
    Set st = New Setup
    
    f_step = 1
    Call StepOne
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Closing
    Set st = Nothing
End Sub

Private Sub UserForm_Terminate()
    ' Clean up
    Set st = Nothing
End Sub

' - Methods

' - Supporting Methods

Private Function StepOne()

    Me.btn_Cancel.Visible = False

    AppendText "We need to evaluate your Outlook folder set " & _
                            "to assure the app can find associated mail messages " & _
                            "and tasks." & vbCrLf & vbCrLf & _
                            "Press Next to continue."
                            
    f_step = 1
                            
End Function

Private Function StepTwo()

    Me.btn_Cancel.Visible = True

    AppendText vbCrLf & vbCrLf & _
                "Press Next to create your initial project list from the Outlook categories " & _
                "otherwise press Cancel to start FocusMe."
                
    f_step = 2

End Function

Private Function FinalStep()

    Me.btn_Cancel.Visible = False
    Me.btn_Execute.Caption = "Get Started"
    
    AppendText vbCrLf & vbCrLf & _
                "Congratulation, you are all set up to run FocusMe for VBA."

    f_step = 3
    
End Function

Private Function ErrorStep()

End Function

' - TextBox

Private Sub AppendText(ByVal strMsg As String)
    If Len(strMsg) > 0 Then _
        Me.txtbx_Instructions.text = Me.txtbx_Instructions.text & strMsg
End Sub

Private Sub ClearText()
    Me.txtbx_Instructions.text = ""
End Sub

' - Status

Private Sub Status(Optional ByVal msg As String = "")
    Me.sb_Status.SimpleText = msg
End Sub
