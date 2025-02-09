VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userformVisualizacao 
   Caption         =   "Visualizar"
   ClientHeight    =   12405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17505
   OleObjectBlob   =   "userformVisualizacao.frx":0000
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
    If toggle_avulso.Caption = "Combos" Then
        Call feedCombos
    Else
        Call feedAvulsos
    End If
End Sub

Private Sub button_clonar_Click()
    
    If toggle_avulso.Caption = "Avulsos" Then
        MsgBox ("Está opção só é valida para os combos")
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

    If toggle_avulso.Caption = "Combos" Then
        Call feedCombos
    Else
        Call feedAvulsos
    End If
End Sub

Private Sub button_gotoCombos_Click()
    Me.Hide
    userformCriacaoCombos.Show
End Sub

Private Sub button_limparCalendario_Click()
    button_calendario.Caption = "Calendario"
    If toggle_avulso.Caption = "Combos" Then
        Call feedCombos
    Else
        Call feedAvulsos
    End If
End Sub



Private Sub list_combos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If toggle_avulso.Caption = "Combos" Then
        userformAlterarCombo.Show
        Call feedCombos
    Else
        userformAlterarAvulso.Show
        Call feedAvulsos
    End If
    
End Sub

Private Sub toggle_avulso_Click()
    If toggle_avulso.Caption = "Combos" Then
        toggle_avulso.Caption = "Avulsos"
        Call feedAvulsos
    Else
        toggle_avulso.Caption = "Combos"
        Call feedCombos
    End If
    
End Sub

Private Sub toggle_descritivo_Click()
    Descritivo.Range("a2:k100").ClearContents
    
    If toggle_descritivo.Caption = "Off" Then
        toggle_descritivo.Caption = "On"
        Call feedDescritivo
    Else
        toggle_descritivo.Caption = "Off"
        toggle_avulso.Caption = "Combos"
        Call feedCombos
    End If
    
End Sub

Private Sub feedDescritivo()
    Dim data As Date
    Dim rg As Range
    
    If button_calendario.Caption = "Calendario" Then
        button_calendario_Click
    End If
    
    data = CDate(button_calendario.Caption)
    
    Set rg = getRangeDescritivo(data)
    
    With list_combos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;40;300;60;60;60;50"
        .ListIndex = 0
    
    End With

End Sub

Private Sub UserForm_Activate()
    If toggle_avulso.Caption = "Combos" Then
        Call feedCombos
    Else
        Call feedAvulsos
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Call feedCombos
End Sub


Private Sub feedCombos()
    Dim data As String
    Dim rg As Range
    
    data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then data = ""
    
    Set rg = getRangeCombos(textbox_itens, data)
    
    With list_combos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;200;100;60;65;80;50;90;90"
        .ListIndex = 0
    
    End With

End Sub
Private Sub feedAvulsos()
    Dim data As String
    Dim rg As Range
    
    data = button_calendario.Caption
    If button_calendario.Caption = "Calendario" Then data = ""
    
    Set rg = getRangeAvulsos(textbox_itens, data)
    
    With list_combos
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnHeads = True
        .ColumnWidths = "0;40;240;45;60;65;80;50;90;90"
        .ListIndex = 0
    
    End With

End Sub
