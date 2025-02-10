VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformDataUso 
   Caption         =   "Data de uso"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "userformDataUso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformDataUso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
    Me.Hide
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    
    button_calendario.Caption = "Calendario"
    
End Sub
