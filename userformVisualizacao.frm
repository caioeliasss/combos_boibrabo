VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformVisualizacao 
   Caption         =   "Visualizar"
   ClientHeight    =   11700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17505
   OleObjectBlob   =   "userformVisualizacao.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userformVisualizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub button_apagarCombo_Click()
    Dim resposta As VbMsgBoxResult
    Dim id As String
    
    pagina = isPage
    
    If pagina = "Avulsos" Then
        resposta = MsgBox("Deseja mesmo deletar esse Avulso?", vbYesNo, "Apagar")
    
        If resposta <> vbYes Then Exit Sub
        
        lista_index = userformVisualizacao.list_avulsos.ListIndex
        
        id = userformVisualizacao.list_avulsos.List(lista_index, 0)
        
        Call deleteDatabase(Avulsos, Avulsos.Range("a1").CurrentRegion, 1, id, 1, 10)
        Call feedAvulsos
    Else
    
    
        resposta = MsgBox("Deseja mesmo deletar esse combo?", vbYesNo, "Apagar")
        
        If resposta <> vbYes Then Exit Sub
        
        lista_index = userformVisualizacao.list_combos.ListIndex
        
        id = userformVisualizacao.list_combos.List(lista_index, 0)
        
        Call deleteDatabase(Combos, Combos.Range("a1").CurrentRegion, 1, id, 1, 10)
        Call deleteDatabase(ProdutosCombo, ProdutosCombo.Range("a1").CurrentRegion, 1, id, 1, 7)
        Call feedCombos
    End If
    
    
    
End Sub

Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
    
    Call limparTextbox
    
    
    Call feedAvulsos
    Call feedCombos
    Call feedDescritivo
    
End Sub

Private Sub limparTextbox()
    textbox_filtroStatus = ""
    textbox_itens = ""
End Sub

Private Function isPage() As String
    pag_index = MultiPage1.Value
    isPage = MultiPage1.Pages(pag_index).Caption
End Function
Private Sub button_clonar_Click()
    
    
    Dim resposta As VbMsgBoxResult
    Dim id As String
    lista_index = userformVisualizacao.list_combos.ListIndex
    If lista_index < 0 Then Exit Sub
    
    resposta = MsgBox("Deseja mesmo clonar esse combo?", vbYesNo, "Clonar")
    
    If resposta <> vbYes Then Exit Sub
    
    
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    Call clonarCombo(id)
    Call feedCombos
    
End Sub

Private Sub button_consultar_Click()
    
    textbox_filtroStatus = ""
    textbox_itens = ""
    
End Sub

Private Sub button_gerarPDF_Click()

        Dim rng As Range
        Dim caminho As String
        Dim nomeArquivo As String
        
        ' Definir a planilha e o intervalo
        Set rng = Descritivo.Range("A1").CurrentRegion
        Set rng = rng.Offset(, 1).Resize(, rng.Columns.Count - 1)
         Descritivo.Visible = xlSheetVisible
        Descritivo.PageSetup.PrintArea = rng.Address

        ' Definir caminho e nome do arquivo
        caminho = ThisWorkbook.Path & "\pdf\"
        If button_calendario.Caption = "Calendario" Then
            nomeArquivo = "Descritivo geral " & Format(Now, "dd-mm-yyyy") & ".pdf"
        Else
            nomeArquivo = "Descritivo " & Format(CDate(button_calendario.Caption), "dd-mm-yyyy") & ".pdf"
        End If
        
        Descritivo.PageSetup.Orientation = xlLandscape
        
        ' Exportar como PDF
        rng.ExportAsFixedFormat Type:=xlTypePDF, _
                                Filename:=caminho & nomeArquivo, _
                                Quality:=xlQualityStandard, _
                                IncludeDocProperties:=True, _
                                IgnorePrintAreas:=False
        
        ' Mensagem de confirma?
        Descritivo.Range("h1").ClearContents
        Descritivo.Range("h2").ClearContents
        Descritivo.Visible = xlSheetHidden
        MsgBox "PDF salvo", vbInformation, "Sucesso"
        

End Sub

Private Sub button_gotoCombos_Click()
    Me.Hide
    userformCriacaoCombos.Show
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
    Call feedAvulsos
    Call feedCombos
    Call feedDescritivo
End Sub



Private Sub button_pagAnterior_Click()
    Me.Hide
    userformCriacaoCombos.Show
End Sub

Private Sub combobox_ordenar_Change()
    Call feedAvulsos
    Call feedCombos
    Call feedDescritivo
End Sub

Private Sub list_combos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    userformAlterarCombo.Show
    Call feedCombos
    
End Sub
Private Sub list_avulsos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    userformAlterarAvulso.Show
    Call feedAvulsos
    
End Sub


Private Sub feedDescritivo()
    Dim Data As Date
    Dim rg As Range
    Dim ordem As Integer
    
    If button_calendario.Caption = "Calendario" Then
            Data = Empty
        Else
            Data = CDate(button_calendario.Caption)
    End If
    
    ordem = combobox_ordenar.ListIndex + 2
    
    Set rg = getRangeDescritivo(Data, textbox_filtroStatus.Text, ordem)
    
    With list_descritivo
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;55;300;30;150;150;50"
        .ListIndex = 0
    
    End With

End Sub

Private Sub list_descritivo_Click()

End Sub

Private Sub MultiPage1_Change()
    limparTextbox
    Call feedCombos
    Call feedAvulsos
    Call feedDescritivo
End Sub

Private Sub textbox_filtroStatus_Change()
    feedCombos
    feedAvulsos
    feedDescritivo
End Sub

Private Sub textbox_itens_Change()
    feedCombos
    feedAvulsos
    feedDescritivo
End Sub

Private Sub UserForm_Activate()
    feedCombos
    feedAvulsos
    feedDescritivo
End Sub

Private Sub UserForm_Initialize()
    feedCombos
    feedAvulsos
    feedDescritivo
    Call feedOrdenar
   MultiPage1.Value = 0
End Sub

Private Sub feedOrdenar()
    
    combobox_ordenar.Clear
    With combobox_ordenar
        .AddItem "Produtos"
        .AddItem "Produto ID"
        .AddItem "Custo"
        .AddItem "Venda"
        .AddItem "Data criacao"
        .AddItem "Data uso"
        .AddItem "Status"
        .AddItem "Intervalo"
    End With
    
    combobox_ordenar.ListIndex = 3

End Sub

Private Sub feedCombos()
    Dim Data As String
    Dim rg As Range
    Dim ordem As Integer
    Dim produto As String
    
    Data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then Data = ""
    
    ordem = combobox_ordenar.ListIndex + 2
    
    produto = RemoverAcentos(textbox_itens.Text)
    
    Set rg = getRangeCombos(produto, Data, ordem, textbox_filtroStatus)
    
    With list_combos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;200;0;60;65;80;50;90;90;100"
        .ListIndex = 0
    
    End With

End Sub
Private Sub feedAvulsos()
    Dim Data As String
    Dim rg As Range
    Dim ordem As Integer
    Dim produto As String
    
    
    Data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then Data = ""
    ordem = combobox_ordenar.ListIndex + 2
    
    produto = RemoverAcentos(textbox_itens.Text)
    
    Set rg = getRangeAvulsos(produto, Data, ordem, textbox_filtroStatus)
    
    With list_avulsos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;40;240;45;60;65;80;50;90;90"
        .ListIndex = 0
    
    End With

End Sub













