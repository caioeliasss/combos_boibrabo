VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformDataUso 
   Caption         =   "Data de uso"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4725
   OleObjectBlob   =   "userformDataUso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformDataUso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public isCanceled As Boolean


Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
End Sub

Private Sub button_cancel_Click()
    isCanceled = True
    Me.Hide
    Call limparCampos
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
    
End Sub

Private Sub button_salvar_Click()
    Me.Hide
    Call limparCampos
End Sub

Private Sub UserForm_Initialize()
    
    isCanceled = False
    button_calendario.Caption = "Calendario"
    Call limparCampos
    
End Sub

Private Sub limparCampos()

    textbox_comentario = ""
    texbox_int = ""
    textbox_status = ""
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

