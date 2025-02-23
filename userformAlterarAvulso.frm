VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformAlterarAvulso 
   Caption         =   "Alterar avulso"
   ClientHeight    =   7035
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
Public id As String



Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
End Sub

Private Sub button_salvar_Click()
Dim Data As String
Dim status As String
Dim obs As String

    lista_index = userformVisualizacao.list_avulsos.ListIndex
    id = userformVisualizacao.list_avulsos.List(lista_index, 0)
    Data = button_calendario.Caption
    If Data = "Calendario" Then Data = ""
    status = textbox_status
    obs = textbox_observacao
    
    If Data = "" Then
        Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8, Data)
    Else
        Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8, CDate(Data))
    End If
    
    peso = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 5)
    
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 9, status)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 10, obs)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 6, textbox_precoVenda)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 4, textbox_peso)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 5, textbox_peso * peso)
    Call updateDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 11, textbox_comentario)
    
    Unload Me
    

End Sub

Private Sub isDateUsed()
    
    
    lista_index = userformVisualizacao.list_avulsos.ListIndex
    
    id = userformVisualizacao.list_avulsos.List(lista_index, 0)
    
    Data = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8)
    If Data = "" Then
    Else
        button_calendario.Caption = Data
    End If
    

End Sub

Private Sub feedProduto()
    
    
    lista_index = userformVisualizacao.list_avulsos.ListIndex
    
    id = userformVisualizacao.list_avulsos.List(lista_index, 0)
    
    dia = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 8)
    If dia = "" Then dia = "Calendario"
    
    button_calendario.Caption = dia
    textbox_custo = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 5)
    textbox_precoVenda = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 6)
    textbox_status = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 9)
    textbox_observacao = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 10)
    textbox_peso = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 4)
    textbox_comentario = consultarDatabase(Avulsos.Range("a1").CurrentRegion, Avulsos, 1, id, 11)
    textbox_lucro = getLucro
    
End Sub

Private Function getLucro()
    getLucro = Round((1 - textbox_custo / textbox_precoVenda) * 100, 1)
End Function

Private Sub textbox_precoVenda_Change()
    textbox_lucro = getLucro
End Sub

Private Sub UserForm_Initialize()
    Call isDateUsed
    Call feedProduto
End Sub




