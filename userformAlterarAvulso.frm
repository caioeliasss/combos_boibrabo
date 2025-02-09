VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformAlterarAvulso 
   Caption         =   "Alterar avulso"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   OleObjectBlob   =   "userformAlterarAvulso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformAlterarAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
End Sub

Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
End Sub

Private Sub button_salvar_Click()
Dim id As String
Dim data As String
Dim status As String
Dim obs As String

    lista_index = userformVisualizacao.list_combos.ListIndex
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    data = button_calendario.Caption
    If data = "Calendario" Then data = ""
    status = textbox_status
    obs = textbox_observacao
    
    If data = "" Then
        Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8, data)
    Else
        Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8, CDate(data))
    End If
    
    peso = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 5)
    
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 9, status)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 10, obs)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 6, textbox_precoVenda)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 4, textbox_peso)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 5, textbox_peso * peso)
    
    
    Unload Me
    

End Sub

Private Sub isDateUsed()
    Dim id As String
    
    
    lista_index = userformVisualizacao.list_combos.ListIndex
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    
    data = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8)
    If data = "" Then
    Else
        button_calendario.Caption = data
    End If
    

End Sub

Private Sub feedProduto()
    Dim id As String
    
    
    lista_index = userformVisualizacao.list_combos.ListIndex
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    
    textbox_precoVenda = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 6)
    textbox_status = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 9)
    textbox_observacao = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 10)
    textbox_peso = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 4)
End Sub

Private Sub UserForm_Initialize()
    Call isDateUsed
    Call feedProduto
End Sub
