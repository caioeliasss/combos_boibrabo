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
    
    If toggle_avulso.Caption = "Avulsos" Then
        resposta = MsgBox("Deseja mesmo deletar esse Avulso?", vbYesNo, "Apagar")
    
        If resposta <> vbYes Then Exit Sub
        
        lista_index = userformVisualizacao.list_combos.ListIndex
        
        id = userformVisualizacao.list_combos.List(lista_index, 0)
        
        Call deleteDatabase(Avulsos, Avulsos.Range("a1").CurrentRegion, 1, id, 1, 10)
        Call feedAvulsos
    Else
    
    
        resposta = MsgBox("Deseja mesmo deletar esse combo?", vbYesNo, "Apagar")
        
        If resposta <> vbYes Then Exit Sub
        
        lista_index = userformVisualizacao.list_combos.ListIndex
        
        id = userformVisualizacao.list_combos.List(lista_index, 0)
        
        Call deleteDatabase(Combos, Combos.Range("a1").CurrentRegion, 1, id, 1, 9)
        Call deleteDatabase(ProdutosCombo, ProdutosCombo.Range("a1").CurrentRegion, 1, id, 1, 7)
        Call feedCombos
    End If
    
    
    
End Sub

Private Sub button_calendario_Click()
    Calendario.Show
    button_calendario.Caption = Calendario.labelDataSelecionada
    
    Call feedAvulsos
    Call feedCombos
    Call feedDescritivo
    
End Sub

Private Sub button_clonar_Click()
    
    If toggle_avulso.Caption = "Avulsos" Then
        MsgBox ("Esta opcao e valida para os combos")
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    Dim id As String
    
    resposta = MsgBox("Deseja mesmo clonar esse combo?", vbYesNo, "Clonar")
    
    If resposta <> vbYes Then Exit Sub
    
    lista_index = userformVisualizacao.list_combos.ListIndex
    
    id = userformVisualizacao.list_combos.List(lista_index, 0)
    Call clonarCombo(id)
    Call feedCombos
    
End Sub

Private Sub button_consultar_Click()
    
    Call feedAvulsos
    Call feedCombos
    Call feedDescritivo
    
End Sub

Private Sub button_gerarPDF_Click()

        Dim rng As Range
        Dim caminho As String
        Dim nomeArquivo As String
        
        Descritivo.Range("h1").Value = "Data de uso"
        Descritivo.Range("h2").Value = Format(Now, "dd/mm/yyyy")
        ' Definir a planilha e o intervalo
        Set rng = Descritivo.Range("A1").CurrentRegion
        Set rng = rng.Offset(, 1).Resize(, rng.Columns.Count - 1)
         Descritivo.Visible = xlSheetVisible
        Descritivo.PageSetup.PrintArea = rng.Address

        ' Definir caminho e nome do arquivo
        caminho = ThisWorkbook.Path & "\"
        nomeArquivo = "Descritivo " & Format(CDate(button_calendario.Caption), "dd-mm-yyyy") & ".pdf"
        
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
        
    End If

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
    Call feedCombos
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
    
    If button_calendario.Caption = "Calendario" Then
            Data = Empty
        Else
            Data = CDate(button_calendario.Caption)
    End If
    
    Set rg = getRangeDescritivo(Data, textbox_filtroStatus.Text)
    
    With list_descritivo
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;40;300;60;60;60;50"
        .ListIndex = 0
    
    End With

End Sub

Private Sub UserForm_Activate()
    feedCombos
    feedAvulsos
    feedDescritivo
    Call feedOrdenar
End Sub

Private Sub UserForm_Initialize()
    feedCombos
    feedAvulsos
    feedDescritivo
    Call feedOrdenar
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
    
    Data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then Data = ""
    
    ordem = combobox_ordenar.ListIndex + 2
    
    Set rg = getRangeCombos(textbox_itens, Data, ordem, textbox_filtroStatus)
    
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
    
    Data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then Data = ""
    
    Set rg = getRangeAvulsos(textbox_itens, Data)
    
    With list_avulsos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;40;240;45;60;65;80;50;90;90"
        .ListIndex = 0
    
    End With

End Sub








