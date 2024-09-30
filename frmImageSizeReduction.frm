VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImageSizeReduction 
   Caption         =   "Select Image Size Reduction"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3675
   OleObjectBlob   =   "frmImageSizeReduction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImageSizeReduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Me.Hide
End Sub

' Only include this if you have a Cancel button named cmdCancel
Private Sub cmdCancel_Click()
    ' Unload the form and exit the macro
    Unload Me
    End
End Sub

Private Sub UserForm_Click()

End Sub
