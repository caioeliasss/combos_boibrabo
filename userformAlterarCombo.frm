VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformAlterarCombo 
   Caption         =   "Modificar"
   ClientHeight    =   8580.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
   OleObjectBlob   =   "userformAlterarCombo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformAlterarCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public custo As Double
Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
End Sub

Private Function totalizadorCusto()
Dim var As Variant

var = ProdutosCombo.Range("AA1").CurrentRegion

If UBound(var) = 1 Then Exit Function

For i = 2 To UBound(var)
    soma = Round(var(i, 7) + soma, 2)
Next i

label_custo = soma

totalizadorCusto = soma

End Function


Private Sub button_editarPeso_Click()
Dim peso As String
Dim id As String
Dim custoNew As String
Dim id_produto As String


lista_index = list_produtosCombo.ListIndex
If lista_index = -1 Then Exit Sub

id = list_produtosCombo.List(lista_index, 0)
id_produto = list_produtosCombo.List(lista_index, 1)

peso = InputBox("Altere o peso", "Peso")
If peso = "" Then Exit Sub
If Not IsNumeric(peso) Then Exit Sub


custo = consultarDatabase(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 2, id_produto, 5)
custoNew = Round(CDbl(peso) * CDbl(custo), 1)

Call updateDatabaseEspecial(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 1, id, 2, id_produto, 6, peso)
Call updateDatabaseEspecial(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 1, id, 2, id_produto, 7, custoNew)

Call feedProdutos
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
End Sub

Private Sub button_salvar_Click()
Dim id As String
Dim Data As String
Dim status As String
Dim obs As String

    lista_index = userformVisualizacao.list_combos.ListIndex
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    Data = button_calendario.Caption
    If Data = "Calendario" Then Data = ""
    status = textbox_status
    obs = textbox_observacao
    
    If Data = "" Then
        Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 7, Data)
    Else
        Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 7, CDate(Data))
    End If
    
    Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 8, status)
    Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 9, obs)
    Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 4, totalizadorCusto)
    Call updateDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 5, textbox_precoVenda)
    
    
    Unload Me
    

End Sub






Private Sub button_salvar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub



Private Sub list_produtosCombo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim peso As String
Dim id As String
Dim custoNew As String
Dim id_produto As String


lista_index = list_produtosCombo.ListIndex
If lista_index = -1 Then Exit Sub

id = list_produtosCombo.List(lista_index, 0)
id_produto = list_produtosCombo.List(lista_index, 1)

peso = InputBox("Altere o peso", "Peso")
If peso = "" Then Exit Sub
If Not IsNumeric(peso) Then Exit Sub


custo = consultarDatabase(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 2, id_produto, 5)
custoNew = Round(CDbl(peso) * CDbl(custo), 1)

Call updateDatabaseEspecial(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 1, id, 2, id_produto, 6, peso)
Call updateDatabaseEspecial(ProdutosCombo.Range("a1").CurrentRegion, ProdutosCombo, 1, id, 2, id_produto, 7, custoNew)

Call feedProdutos
End Sub



Private Sub textbox_precoVenda_Change()
    On Error Resume Next
    label_porcentagem = Round((textbox_precoVenda.Value / custo * 100) - 100, 1)
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()

    Call feedProdutos
    Call isDateUsed
    
End Sub

Private Sub isDateUsed()
    Dim id As String
    
    
    lista_index = userformVisualizacao.list_combos.ListIndex
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    
    Data = consultarDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 7)
    If Data = "" Then
    Else
        userformAlterarCombo.button_calendario.Caption = Data
    End If
    

End Sub

Private Sub feedProdutos()
    Dim rg As Range
    Dim id As String
    
    
    lista_index = userformVisualizacao.list_combos.ListIndex
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    Set rg = getRangeComboProdutos(id)
    
    With list_produtosCombo
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;60;250;50;60;50"
        .ListIndex = 0
    End With
    
    custo = consultarDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 4)
    
    textbox_precoVenda = consultarDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 5)
    textbox_status = consultarDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 8)
    textbox_observacao = consultarDatabase(Combos.Range("a1").CurrentRegion, Combos, 1, id, 9)
    label_porcentagem = Round((textbox_precoVenda.Value / custo) - 1, 1)
    
    

End Sub


