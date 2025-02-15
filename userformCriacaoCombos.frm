VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformCriacaoCombos 
   Caption         =   "Combos BOIBRABO"
   ClientHeight    =   11835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19860
   OleObjectBlob   =   "userformCriacaoCombos.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformCriacaoCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Public DataSelecionada As Date





Private Sub button_editarPeso_Click()

Dim id As String

lista_index = listCombos.ListIndex

id = listCombos.List(lista_index, 0)

Call deleteDatabase(Produtos, Produtos.Range("u1").CurrentRegion, 1, id, 21, 28)
Call feedHeaderNew
Call totalizadorCusto
Call totalizadorVendaFora

End Sub



Private Sub button_gotoVisualizacao_Click()
    Me.Hide
    userformVisualizacao.Show
End Sub

Private Sub button_proximaPag_Click()
    Me.Hide
    userformVisualizacao.Show
End Sub

Private Sub button_salvar_Click()
Dim combo_id As String
Dim produtos_combo As Range
Dim var_produtosCombo As Variant
Dim status As String
Dim observacao As String
Dim avulsoRange As Range
Dim varAvulsos As Variant


    
    userformDataUso.Show
    
    If userformDataUso.isCanceled = True Then Exit Sub
    
    data_uso = userformDataUso.button_calendario.Caption
    If data_uso = "Calendario" Then data_uso = ""
    If data_uso <> "" Then data_uso = CDate(data_uso)
    
    status = userformDataUso.textbox_status
    observacao = userformDataUso.texbox_int
    comentario = userformDataUso.textbox_comentario
    

If listCombos.ListCount = 1 Then

    avulso_id = WorksheetFunction.RandBetween(111111111, 999999999)
    
    last_row_avulso = Avulsos.Range("a1").CurrentRegion.Rows.Count + 1
    
    Set avulsoRange = Produtos.Range("u1").CurrentRegion
    varAvulsos = avulsoRange.Offset(1).Resize(avulsoRange.Rows.Count - 1)


   
    Avulsos.Cells(last_row_avulso, 1) = avulso_id
    Avulsos.Cells(last_row_avulso, 2) = varAvulsos(1, 1)
    Avulsos.Cells(last_row_avulso, 3) = varAvulsos(1, 2)
    Avulsos.Cells(last_row_avulso, 4) = varAvulsos(1, 5)
    Avulsos.Cells(last_row_avulso, 5) = CDbl(varAvulsos(1, 6))
    Avulsos.Cells(last_row_avulso, 6) = CDbl(textbox_venda)
    Avulsos.Cells(last_row_avulso, 7) = Date
    Avulsos.Cells(last_row_avulso, 8) = data_uso
    Avulsos.Cells(last_row_avulso, 9) = status
    Avulsos.Cells(last_row_avulso, 10) = observacao
    Avulsos.Cells(last_row_avulso, 11) = comentario
    
Else

    combo_id = WorksheetFunction.RandBetween(111111111, 999999999)
    
    last_row_produtoCombo = ProdutosCombo.Range("a1").CurrentRegion.Rows.Count + 1
    last_row_combos = Combos.Range("a1").CurrentRegion.Rows.Count + 1
    
    Set produtos_combo = Produtos.Range("u1").CurrentRegion
    produtos_combo.Sort Key1:=produtos_combo.Columns(8), Order1:=xlDescending, Header:=xlNo

    var_produtosCombo = produtos_combo.Offset(1).Resize(produtos_combo.Rows.Count - 1)
    
    ProdutosCombo.Range(Cells(last_row_produtoCombo, 2).Address, Cells(last_row_produtoCombo + UBound(var_produtosCombo) - 1, 7).Address).Value = var_produtosCombo
    ProdutosCombo.Range(Cells(last_row_produtoCombo, 1).Address, Cells(last_row_produtoCombo + UBound(var_produtosCombo) - 1, 1).Address).Value = combo_id
    
    '--------- Combos
    
    For i = 1 To UBound(var_produtosCombo)
        virgula = ", "
        If i = UBound(var_produtosCombo) Then virgula = ""
        
        lista_produtos = lista_produtos & var_produtosCombo(i, 2) & virgula
        lista_produto_id = lista_produto_id & var_produtosCombo(i, 1) & virgula
    
    Next i
    
    Combos.Cells(last_row_combos, 1) = combo_id
    Combos.Cells(last_row_combos, 2) = lista_produtos
    Combos.Cells(last_row_combos, 3) = lista_produto_id
    Combos.Cells(last_row_combos, 4) = CDbl(label_custo)
    Combos.Cells(last_row_combos, 5) = CDbl(textbox_venda)
    Combos.Cells(last_row_combos, 6) = Date
    Combos.Cells(last_row_combos, 7) = data_uso
    Combos.Cells(last_row_combos, 8) = status
    Combos.Cells(last_row_combos, 9) = observacao
    Combos.Cells(last_row_combos, 10) = comentario

End If

Produtos.Range("u2:ab1000").ClearContents


End Sub

Private Sub listCombos_Click()
    Call nameProduto(listCombos)
End Sub

Private Sub listCombos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim peso As String
Dim id As String
Dim custoNew As String


lista_index = listCombos.ListIndex
If lista_index = -1 Then Exit Sub

id = listCombos.List(lista_index, 0)

peso = InputBox("Insira o peso", "Peso")
If peso = "" Then Exit Sub
If Not IsNumeric(peso) Then Exit Sub



custo = consultarDatabase(Produtos.Range("u1").CurrentRegion, Produtos, 1, id, 4)
custoNew = Round(CDbl(peso) * CDbl(custo), 1)

venda = consultarDatabase(Produtos.Range("u1").CurrentRegion, Produtos, 1, id, 7)
vendaNew = Round(CDbl(peso) * CDbl(venda), 1)

Call updateDatabase(Produtos.Range("u1").CurrentRegion, Produtos, 1, id, 25, peso)
Call updateDatabase(Produtos.Range("u1").CurrentRegion, Produtos, 1, id, 26, custoNew)
Call updateDatabase(Produtos.Range("u1").CurrentRegion, Produtos, 1, id, 28, vendaNew)
Call updateDatabase(Produtos.Range("a1").CurrentRegion, Produtos, 1, id, 14, CDbl(peso))

Call totalizadorCusto
Call totalizadorVendaFora
Call feedHeaderNew
Call feedProdutos

End Sub

Private Sub listProdutos_click()
    Call nameProduto(listProdutos)
End Sub

Private Sub nameProduto(lista As Object)
    index_atual = lista.ListIndex
    If index_atual = -1 Then Exit Sub
    label_produto.Caption = lista.List(index_atual, 1)
End Sub

Private Sub listProdutos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim lista_index_produto As Integer

    lista_index_produto = listProdutos.ListIndex
    
    ' Verifica se h?m item selecionado
    If lista_index_produto = -1 Then Exit Sub
    
    last_row = Produtos.Range("u1").CurrentRegion.Rows.Count + 1
    
    Produtos.Cells(last_row, 21) = listProdutos.List(lista_index_produto, 0)
    Produtos.Cells(last_row, 22) = listProdutos.List(lista_index_produto, 1)
    Produtos.Cells(last_row, 23) = listProdutos.List(lista_index_produto, 2)
    Produtos.Cells(last_row, 24) = Round(listProdutos.List(lista_index_produto, 3), 1)
    Produtos.Cells(last_row, 25) = listProdutos.List(lista_index_produto, 13)
    Produtos.Cells(last_row, 26) = Round(listProdutos.List(lista_index_produto, 3), 1) * listProdutos.List(lista_index_produto, 13)
    Produtos.Cells(last_row, 27) = Round(listProdutos.List(lista_index_produto, 5), 1)
    Produtos.Cells(last_row, 28) = Round(listProdutos.List(lista_index_produto, 5), 1) * listProdutos.List(lista_index_produto, 13)
    
    Call feedHeaderNew
    Call totalizadorCusto
    Call totalizadorVendaFora
    Call calcularDesconto
    
End Sub

Private Sub calcularVenda()
On Error Resume Next

porcentagem = 1 - ((textbox_porcentagem) * 0.01)

textbox_venda = label_custo / porcentagem
textbox_venda = Round(textbox_venda, 2)

On Error GoTo 0
End Sub

Private Sub totalizadorVendaFora()
Dim var As Variant

var = Produtos.Range("u1").CurrentRegion

If UBound(var) = 1 Then Exit Sub

For i = 2 To UBound(var)
    soma = var(i, 8) + soma
Next i

label_venda_foracombo = Round(soma, 1)

End Sub


Private Sub totalizadorCusto()
Dim var As Variant

var = Produtos.Range("u1").CurrentRegion

If UBound(var) = 1 Then Exit Sub

For i = 2 To UBound(var)
    soma = var(i, 6) + soma
Next i

label_custo = Round(soma, 1)

Call calcularVenda

End Sub


Private Sub button_consultar_Click()

    textbox_idproduto = ""
    textbox_produto = ""
    Call feedProdutos

End Sub

Private Sub button_favoritar_Click()
Dim id As String

lista_index = listProdutos.ListIndex

id = listProdutos.List(lista_index, 0)

favorito = consultarDatabase(Produtos.Range("a1").CurrentRegion, Produtos, 1, id, 13)

If favorito = "" Then
    Call updateDatabase(Produtos.Range("a1").CurrentRegion, Produtos, 1, id, 13, "sim")
Else
    Call updateDatabase(Produtos.Range("a1").CurrentRegion, Produtos, 1, id, 13, "")
End If

Call feedProdutos

If listProdutos.ListCount = lista_index Then lista_index = lista_index - 1
listProdutos.ListIndex = lista_index

End Sub


Private Sub button_verFavoritos_Click()

If placeholder_favorito = "sim" Then
    placeholder_favorito = ""
Else
    placeholder_favorito = "sim"
End If

Call feedProdutos

End Sub



Private Sub textbox_idproduto_Change()
    Call feedProdutos
End Sub

Private Sub textbox_porcentagem_Change()
    Call calcularVenda
End Sub

Private Sub textbox_produto_Change()
    Call feedProdutos
End Sub

Private Sub textbox_venda_Change()
    On Error Resume Next ' Ignora erros e continua a execu?

    If Val(textbox_venda) = 0 Then Exit Sub

    label_porcentagem = Val((1 - (Val(label_custo) / Val(textbox_venda))) * 100)
    label_lucro = Val(textbox_venda) - Val(label_custo)
    
    Call calcularDesconto

    On Error GoTo 0 ' Desativa o tratamento de erro para evitar ignorar outros erros inesperados
End Sub

Private Sub calcularDesconto()

    textbox_desconto = Round(label_venda_foracombo - textbox_venda, 2)

End Sub

Private Sub UserForm_Initialize()


Produtos.Range("u2", "ab100").Clear

Call feedProdutos
Call feedHeaderNew
Call feedValues

End Sub

Private Sub feedValues()

textbox_porcentagem = 30

End Sub

Private Sub feedHeaderNew()
Dim rg As Range


    
Set rg = Produtos.Range("u1").CurrentRegion

If rg.Rows.Count - 1 = 0 Then
    Set rg = rg.Offset(1).Resize(1)
Else
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
End If

' Configura o ListBox do cabe?ho
    listCombos.ColumnCount = rg.Columns.Count
    listCombos.RowSource = rg.Address(external:=True)
    listCombos.ColumnHeads = True
    listCombos.ColumnWidths = "35;150;30;65;35;60;45;45"
End Sub


Private Sub feedProdutos()
Dim rg As Range

produto = RemoverAcentos(textbox_produto.Text)

Set rg = getRangeProdutos(textbox_produto, textbox_idproduto, placeholder_favorito)

With listProdutos
    .RowSource = rg.Address(external:=True)
    .ColumnCount = rg.Columns.Count
    .ColumnHeads = True
    .ColumnWidths = "40;200;35;50;0;0;0;0;0;0;0;50;50;0"
    .ListIndex = 0
End With




End Sub










