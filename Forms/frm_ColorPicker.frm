VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ColorPicker 
   Caption         =   "Select Color"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3180
   OleObjectBlob   =   "frm_ColorPicker.frx":0000
End
Attribute VB_Name = "frm_ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' - - Fields

Dim f_Selection As Integer
Dim f_Picture As IPictureDisp

' - - Properties

''' Selected Color
Public Property Let Selection(ByVal iVal As Integer)
    f_Selection = iVal
End Property
Public Property Get Selection() As Integer
    Selection = f_Selection
End Property

''' Selected Picture
Public Property Get SelectedPicture() As IPictureDisp
    Set SelectedPicture = f_Picture
End Property

' - - Event Handlers

Private Sub Image1_Click()
    ' Not used
End Sub

Private Sub Image2_Click()
    ' Red
    f_Selection = 1
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image3_Click()
    ' Orange
    f_Selection = 2
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image4_Click()
    ' Peach
    f_Selection = 3
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image5_Click()
    f_Selection = 4
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image6_Click()
    f_Selection = 5
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image7_Click()
    f_Selection = 6
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image8_Click()
    f_Selection = 7
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image9_Click()
    f_Selection = 8
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image10_Click()
    f_Selection = 9
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image11_Click()
    f_Selection = 10
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image12_Click()
    f_Selection = 11
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image13_Click()
    f_Selection = 12
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image14_Click()
    f_Selection = 13
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image15_Click()
    f_Selection = 14
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image16_Click()
    f_Selection = 15
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image17_Click()
    f_Selection = 16
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image18_Click()
    f_Selection = 17
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image19_Click()
    f_Selection = 18
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image20_Click()
    f_Selection = 19
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image21_Click()
    f_Selection = 20
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image22_Click()
    f_Selection = 21
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image23_Click()
    f_Selection = 22
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image24_Click()
    f_Selection = 23
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image25_Click()
    f_Selection = 24
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub Image26_Click()
    f_Selection = 25
    img_Selection.Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
End Sub

Private Sub btn_Clear_Click()
    ' No Category
    img_Selection.Picture = Image1.Picture
    f_Selection = 1
End Sub

Private Sub btn_Save_Click()
    Set f_Picture = imglst_Colors.ListImages(f_Selection + 1).Picture
    Me.Hide
End Sub

' - - Constructor

Private Sub UserForm_Initialize()

    ' Create imglst from Palette Tab
    imglst_Colors.ListImages.Add 1, "None", Image1.Picture
    imglst_Colors.ListImages.Add 2, "Red", Image2.Picture       ' olCategoryColorRed = 1
    imglst_Colors.ListImages.Add 3, "Orange", Image3.Picture
    imglst_Colors.ListImages.Add 4, "Peach", Image4.Picture
    imglst_Colors.ListImages.Add 5, "Yellow", Image5.Picture    ' olCategoryColorYellow = 4
    imglst_Colors.ListImages.Add 6, "Green", Image6.Picture
    imglst_Colors.ListImages.Add 7, "Teal", Image7.Picture
    imglst_Colors.ListImages.Add 8, "Olive", Image8.Picture
    imglst_Colors.ListImages.Add 9, "Blue", Image9.Picture
    imglst_Colors.ListImages.Add 10, "Purple", Image10.Picture
    imglst_Colors.ListImages.Add 11, "Maroon", Image11.Picture  ' olCategoryColorMaroon = 10
    imglst_Colors.ListImages.Add 12, "Steel", Image12.Picture
    imglst_Colors.ListImages.Add 13, "Dark Steel", Image13.Picture
    imglst_Colors.ListImages.Add 14, "Gray", Image14.Picture
    imglst_Colors.ListImages.Add 15, "Dark Gray", Image15.Picture
    imglst_Colors.ListImages.Add 16, "Black", Image16.Picture
    imglst_Colors.ListImages.Add 17, "Dark Red", Image17.Picture
    imglst_Colors.ListImages.Add 18, "Dark Orange", Image18.Picture
    imglst_Colors.ListImages.Add 19, "Dark Peach", Image19.Picture
    imglst_Colors.ListImages.Add 20, "Dark Yellow", Image20.Picture
    imglst_Colors.ListImages.Add 21, "Dark Green", Image21.Picture
    imglst_Colors.ListImages.Add 22, "Dark Teal", Image22.Picture
    imglst_Colors.ListImages.Add 23, "Dark Olive", Image23.Picture
    imglst_Colors.ListImages.Add 24, "Dark Blue", Image24.Picture
    imglst_Colors.ListImages.Add 25, "Dark Purple", Image25.Picture
    imglst_Colors.ListImages.Add 26, "Dark Maroon", Image26.Picture ' olCategoryColorDarkMaroon = 25
    
    Set f_Picture = Nothing

End Sub

' - - Methods

