Attribute VB_Name = "modRange"
Public Sub setHeaders()
    
    Produtos.Range("w1").Value = "Un."
    Descritivo.Range("d1").Value = "Un."
    Combos.Range("j1").Value = "Comentarios"
    Combos.Range("aj1").Value = "Comentarios"
    
    Avulsos.Range("k1").Value = "Comentarios"
    Avulsos.Range("ak1").Value = "Comentarios"
    
    Produtos.Visible = xlSheetHidden
    Combos.Visible = xlSheetHidden
    ProdutosCombo.Visible = xlSheetHidden
    Avulsos.Visible = xlSheetHidden
    Descritivo.Visible = xlSheetHidden

    meuId = consultarDatabase(Log.Range("a1").CurrentRegion, Log, 1, "meu_id", 2)
        If meuId = "" Then
            Call createDocument(Log, Array("meu_id", 1))
        End If

End Sub

Public Sub clonarCombo(id As String)
Dim var As Variant
Dim var_prod As Variant

var = Combos.Range("a1").CurrentRegion
var_prod = ProdutosCombo.Range("a1").CurrentRegion

last_row = UBound(var)

For i = 1 To last_row
    If var(i, 1) = id Then
        For col = 1 To UBound(var, 2)
            Combos.Cells(last_row + 1, col) = var(i, col)
        Next col
        
        novo_id = WorksheetFunction.RandBetween(11111111, 99999999)
        Combos.Cells(last_row + 1, 1) = novo_id
        Combos.Cells(last_row + 1, 6) = Date
        Combos.Cells(last_row + 1, 7) = ""
        Combos.Cells(last_row + 1, 8) = ""
    End If
Next i


last_row = UBound(var_prod)
count_ = 1

For i = 1 To last_row
    If var_prod(i, 1) = id Then
        For col = 1 To UBound(var_prod, 2)
            ProdutosCombo.Cells(last_row + count_, col) = var_prod(i, col)
        Next col
        ProdutosCombo.Cells(last_row + count_, 1) = novo_id
        count_ = count_ + 1
    End If
Next i


End Sub

Public Function getRangeDescritivo(dia As Variant, filtro_status As String, ordem As Integer) As Range
Dim rg As Range
Dim var As Variant
Dim comboVar As Variant
Dim avulsoVar As Variant
Dim var2 As Variant
Dim filteredVar As Variant
ReDim filteredVar(1 To 1000, 1 To 15)
ReDim comboVar(1 To 500, 1 To 15)
ReDim avulsoVar(1 To 500, 1 To 11)

Descritivo.Range("a2:h1000").ClearContents

Set rg = Combos.Range("a1").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
rg.Sort Key1:=rg.Columns(ordem), Order1:=xlDescending, Header:=xlNo
var = rg
count_ = 1

If filtro_status = "" Then IsEmpty_filtro_status = True
If dia = Empty Then IsEmpty_dataUso = True

For i = 1 To UBound(var)
    If IsEmpty_filtro_status Then filtro_status = var(i, 8)
    
    If IsEmpty_dataUso Then dia = var(i, 7)

    If IsEmpty_filtro_status Then
        filtro = True
    Else
        filtro = InStr(1, UCase(var(i, 8)), UCase(filtro_status), vbTextCompare) > 0
    End If
    
    If var(i, 7) = dia And filtro Then
        For col = 1 To UBound(var, 2)
            comboVar(count_, col) = var(i, col)
        Next col
        count_ = count_ + 1
    End If
Next i

var = ProdutosCombo.Range("a1").CurrentRegion
count_ = 1
count_2 = 1

For j = 1 To UBound(comboVar)
    If comboVar(j, 1) <> "" Then
        filteredVar(count_, 1) = String(100, "-")
        filteredVar(count_, 2) = comboVar(j, 1)
        filteredVar(count_, 3) = "COMBO " & count_2 & " | Valor: R$" & comboVar(j, 5) & " | Data: " & comboVar(j, 7)
        filteredVar(count_, 4) = String(100, "-")
        filteredVar(count_, 5) = comboVar(j, 8)
        filteredVar(count_, 6) = comboVar(j, 9)
        filteredVar(count_, 7) = comboVar(j, 10)
        count_ = count_ + 1
        For i = 1 To UBound(var)
            If var(i, 1) = comboVar(j, 1) Then
                For col = 1 To UBound(var, 2)
                    If col = 7 Or col = 5 Then
                    Else
                    
                        filteredVar(count_, col) = var(i, col)
                    End If
                Next col
                count_ = count_ + 1
            End If
        Next i
        
        count_2 = count_2 + 1
    Else: Exit For
    End If
    
    For col = 1 To UBound(var, 2)
        filteredVar(count_, col) = String(100, "-")
    Next col
    'filteredVar(count_, 1) = "-"
    count_ = count_ + 1
Next j

var = Avulsos.Range("a1").CurrentRegion

count_3 = 1

For i = 2 To UBound(var)
    If IsEmpty_dataUso Then dia = CDate(var(i, 8))
    If var(i, 8) = dia Then
        For col = 1 To UBound(var, 2)
            avulsoVar(count_3, col) = var(i, col)
        Next col
        count_3 = count_3 + 1
    End If
Next i

count_2 = 1

For i = 1 To UBound(avulsoVar)
    If avulsoVar(i, 1) <> "" Then
        'filteredVar(count_, 1) = String(100, "-")
        'filteredVar(count_, 2) = String(100, "-")
        filteredVar(count_, 3) = "AVULSO " & count_2 & " | Valor: R$" & avulsoVar(i, 6)
        'filteredVar(count_, 4) = String(100, "-")
        filteredVar(count_, 5) = avulsoVar(i, 9)
        filteredVar(count_, 6) = avulsoVar(i, 10)
        filteredVar(count_, 7) = avulsoVar(i, 11)
        count_ = count_ + 1
        count_2 = count_2 + 1
        For col = 1 To 3
            filteredVar(count_, col) = avulsoVar(i, col)
        Next col
        filteredVar(count_, 6) = avulsoVar(i, 4)
        count_ = count_ + 1

        For col = 1 To 7
            filteredVar(count_, col) = String(100, "-")
        Next col
        count_ = count_ + 1
    End If
Next i



If filteredVar(1, 1) = "" Then
    Descritivo.Range("A2").Value = "Nada encontrado"
Else
    Descritivo.Range("A2", Cells(count_, 15).Address).Value = filteredVar
End If

Set rg = Descritivo.Range("A2").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)

Set getRangeDescritivo = rg



End Function
Public Function getRangeAvulsos(pesquisa_produto As String, dataUso As String, ordem As Integer, filtro_status As String) As Range
Dim rg As Range
Dim var As Variant
Dim filteredVar As Variant
ReDim filteredVar(1 To 1000, 1 To 15)
Dim frase As String

Avulsos.Range("aa2:az1000").ClearContents

Set rg = Avulsos.Range("a1").CurrentRegion
If rg.Rows.Count = 1 Then
    Set rg = rg.Offset(1).Resize(rg.Rows.Count)
Else
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
End If

var = rg

If pesquisa_produto = "" Then IsEmpty_pesquisaProduto = True
If dataUso = "" Then IsEmpty_dataUso = True
If filtro_status = "" Then IsEmpty_filtro_status = True
count_ = 1

For i = 1 To UBound(var)

    If IsEmpty_filtro_status Then
        filtro_status = var(i, 9)
    End If
    If IsEmpty_pesquisaProduto Then
        pesquisa_produto = var(i, 3)
    End If
    If IsEmpty_dataUso Then
        dataUso = var(i, 8)
    End If
    If IsEmpty_filtro_status Then
        filtro = True
    Else
        filtro = InStr(1, var(i, 8), UCase(filtro_status), vbTextCompare) > 0
    End If
    
    frase = var(i, 3)
    frase = RemoverAcentos(frase)
    
    If IsEmpty_pesquisaProduto Then
        frase_pesquisa = True
    Else
        frase_pesquisa = InStr(1, frase, UCase(pesquisa_produto), vbTextCompare) > 0
    End If
    
    If frase_pesquisa And var(i, 8) = dataUso And filtro Then
        For col = 1 To UBound(var, 2)
            filteredVar(count_, col) = var(i, col)
        Next col
    count_ = count_ + 1
    End If

Next i

If filteredVar(1, 1) = "" Then
    Avulsos.Range("AA2").Value = "Nada encontrado"
Else
    Avulsos.Range("AA2", Cells(count_, 37).Address).Value = filteredVar
End If

Set rg = Avulsos.Range("AA2").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
rg.Sort Key1:=rg.Columns(ordem), Order1:=xlDescending, Header:=xlNo

Set getRangeAvulsos = rg

End Function
Public Function getRangeCombos(pesquisa_produto As String, dataUso As String, ordem As Integer, filtro_status As String) As Range
Dim rg As Range
Dim var As Variant
Dim filteredVar As Variant
ReDim filteredVar(1 To 1000, 1 To 15)
Dim filtro As Boolean
Dim frase As String

Combos.Range("aa2:az1000").ClearContents

Set rg = Combos.Range("a1").CurrentRegion
If rg.Rows.Count = 1 Then
    Set rg = rg.Offset(1).Resize(rg.Rows.Count)
Else
    Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
End If

var = rg

If pesquisa_produto = "" Then IsEmpty_pesquisaProduto = True
If dataUso = "" Then IsEmpty_dataUso = True
If filtro_status = "" Then IsEmpty_filtro_status = True

count_ = 1

For i = 1 To UBound(var)

    If IsEmpty_filtro_status Then
        filtro_status = var(i, 8)
    End If

    If IsEmpty_pesquisaProduto Then
        pesquisa_produto = var(i, 2)
    End If
    If IsEmpty_dataUso Then
        dataUso = var(i, 7)
    End If
    
    If IsEmpty_filtro_status Then
        filtro = True
    Else
        filtro = InStr(1, var(i, 8), UCase(filtro_status), vbTextCompare) > 0
    End If
    
    frase = var(i, 2)
    frase = RemoverAcentos(frase)
    
    If IsEmpty_pesquisaProduto Then
        frase_pesquisa = True
    Else
        frase_pesquisa = InStr(1, frase, UCase(pesquisa_produto), vbTextCompare) > 0
    End If
    
    If frase_pesquisa And var(i, 7) = dataUso And filtro Then
        For col = 1 To UBound(var, 2)
            filteredVar(count_, col) = var(i, col)
        Next col
    count_ = count_ + 1
    End If

Next i

If filteredVar(1, 1) = "" Then
    Combos.Range("AA2").Value = "Nada encontrado"
Else
    Combos.Range("AA2", Cells(count_, 36).Address).Value = filteredVar
End If

Set rg = Combos.Range("AA2").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
rg.Sort Key1:=rg.Columns(ordem), Order1:=xlDescending, Header:=xlNo

Set getRangeCombos = rg


End Function

Public Function getRangeComboProdutos(id As String) As Range
Dim rg As Range
Dim var As Variant
Dim filteredVar As Variant
ReDim filteredVar(1 To 1000, 1 To 15)

ProdutosCombo.Range("aa2:az1000").ClearContents

Set rg = ProdutosCombo.Range("a1").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)

var = rg
count_ = 1


For i = 1 To UBound(var)
    
   If var(i, 1) = id Then
        For col = 1 To UBound(var, 2)
            filteredVar(count_, col) = var(i, col)
        Next col
        count_ = count_ + 1
    End If

Next i

If filteredVar(1, 1) = "" Then
    ProdutosCombo.Range("AA2").Value = "Nada encontrado"
Else
    ProdutosCombo.Range("AA2", Cells(count_, 35).Address).Value = filteredVar
End If

Set rg = ProdutosCombo.Range("AA2").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
rg.Sort Key1:=rg.Columns(5), Order1:=xlAscending, Header:=xlNo

Set getRangeComboProdutos = rg


End Function

Public Function getRangeProdutos(pesquisa_nome As String, pesquisa_id As String, favorito As String) As Range
Dim rg As Range
Dim var As Variant
Dim filteredVar As Variant
ReDim filteredVar(1 To 1000, 1 To 15)
Dim frase As String

Call setHeaders
Call apagarVestigios

Set rg = Produtos.Range("a1").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)

var = rg
count_ = 1


If favorito = "" Then isEmpty_favorito = True
If pesquisa_nome = "" Then isEmpty_pesquisa_nome = True
If pesquisa_id = "" Then isEmpty_pesquisa_id = True

For i = 1 To UBound(var)
    
    If isEmpty_pesquisa_nome Then
        pesquisa_nome = var(i, 2)
    End If
    If isEmpty_pesquisa_id Then
        pesquisa_id = var(i, 1)
    End If
    If isEmpty_favorito Then
        favorito = var(i, 13)
    End If
    
    frase = var(i, 2)
    
    frase = RemoverAcentos(frase)
    
    If InStr(1, frase, UCase(pesquisa_nome), vbTextCompare) > 0 And pesquisa_id = var(i, 1) And var(i, 13) = favorito Then
         For col = 1 To UBound(var, 2)
             filteredVar(count_, col) = var(i, col)
         Next col
         count_ = count_ + 1
     End If

Next i

If filteredVar(1, 1) = "" Then
    Produtos.Range("AJ2").Value = "Nada encontrado"
Else
    Produtos.Range("AJ2", Cells(count_, 49).Address).Value = filteredVar
End If

Set rg = Produtos.Range("AJ2").CurrentRegion
Set rg = rg.Offset(1).Resize(rg.Rows.Count - 1)
rg.Sort Key1:=rg.Columns(4), Order1:=xlAscending, Header:=xlNo

Set getRangeProdutos = rg


End Function

Public Sub updateDatabase(rg As Range, Table As Worksheet, id_coluna As Integer, id As String, colunaUpdate As Integer, valueUpdate As Variant)
Dim var As Variant

var = rg
If VarType(valueUpdate) = vbDate Then
    valueUpdate = CDate(valueUpdate)
End If

For i = 1 To UBound(var)

    If var(i, id_coluna) = id Then
        Table.Cells(i, colunaUpdate) = valueUpdate
    End If
    


Next i

End Sub
Public Sub updateDatabaseEspecial(rg As Range, Table As Worksheet, id_coluna As Integer, id As String, id_coluna2 As Integer, id_2 As String, colunaUpdate As Integer, valueUpdate As String)
Dim var As Variant

var = rg

For i = 1 To UBound(var)

    If var(i, id_coluna) = id And var(i, id_coluna2) = id_2 Then
        Table.Cells(i, colunaUpdate) = valueUpdate
    End If
    


Next i

End Sub

Public Function consultarDatabase(rg As Range, Table As Worksheet, id_coluna As Integer, id As String, colunaConsulta As Integer)
Dim var As Variant

var = rg

For i = 1 To UBound(var)

    If var(i, id_coluna) = id Then
        consultarDatabase = var(i, colunaConsulta)
        Exit Function
    End If
    
Next i

consultarDatabase = ""

End Function

Public Sub deleteDatabase(Table As Worksheet, rangeDatabase As Range, id_coluna As Integer, id As String, col_inicio, col_final)
Dim var As Variant

var = rangeDatabase

For i = UBound(var) To 1 Step -1

    If var(i, id_coluna) = id Then
        Table.Range(Cells(i, col_inicio).Address, Cells(i, col_final).Address).Delete Shift:=xlUp
    End If
    
Next i

End Sub


public Function RemoverAcentos(texto As String) As String
    Dim resultado As String
    resultado = texto

    ' Substituir letras maiúsculas
    resultado = Replace(resultado, ChrW(193), "A") ' Á
    resultado = Replace(resultado, ChrW(192), "A") ' À
    resultado = Replace(resultado, ChrW(194), "A") ' Â
    resultado = Replace(resultado, ChrW(195), "A") ' Ã
    resultado = Replace(resultado, ChrW(196), "A") ' Ä

    resultado = Replace(resultado, ChrW(201), "E") ' É
    resultado = Replace(resultado, ChrW(200), "E") ' È
    resultado = Replace(resultado, ChrW(202), "E") ' Ê
    resultado = Replace(resultado, ChrW(203), "E") ' Ë

    resultado = Replace(resultado, ChrW(205), "I") ' Í
    resultado = Replace(resultado, ChrW(204), "I") ' Ì
    resultado = Replace(resultado, ChrW(206), "I") ' Î
    resultado = Replace(resultado, ChrW(207), "I") ' Ï

    resultado = Replace(resultado, ChrW(211), "O") ' Ó
    resultado = Replace(resultado, ChrW(210), "O") ' Ò
    resultado = Replace(resultado, ChrW(212), "O") ' Ô
    resultado = Replace(resultado, ChrW(213), "O") ' Õ
    resultado = Replace(resultado, ChrW(214), "O") ' Ö

    resultado = Replace(resultado, ChrW(218), "U") ' Ú
    resultado = Replace(resultado, ChrW(217), "U") ' Ù
    resultado = Replace(resultado, ChrW(219), "U") ' Û
    resultado = Replace(resultado, ChrW(220), "U") ' Ü

    resultado = Replace(resultado, ChrW(199), "C") ' Ç

    ' Substituir letras minúsculas
    resultado = Replace(resultado, ChrW(225), "a") ' á
    resultado = Replace(resultado, ChrW(224), "a") ' à
    resultado = Replace(resultado, ChrW(226), "a") ' â
    resultado = Replace(resultado, ChrW(227), "a") ' ã
    resultado = Replace(resultado, ChrW(228), "a") ' ä

    resultado = Replace(resultado, ChrW(233), "e") ' é
    resultado = Replace(resultado, ChrW(232), "e") ' è
    resultado = Replace(resultado, ChrW(234), "e") ' ê
    resultado = Replace(resultado, ChrW(235), "e") ' ë

    resultado = Replace(resultado, ChrW(237), "i") ' í
    resultado = Replace(resultado, ChrW(236), "i") ' ì
    resultado = Replace(resultado, ChrW(238), "i") ' î
    resultado = Replace(resultado, ChrW(239), "i") ' ï

    resultado = Replace(resultado, ChrW(243), "o") ' ó
    resultado = Replace(resultado, ChrW(242), "o") ' ò
    resultado = Replace(resultado, ChrW(244), "o") ' ô
    resultado = Replace(resultado, ChrW(245), "o") ' õ
    resultado = Replace(resultado, ChrW(246), "o") ' ö

    resultado = Replace(resultado, ChrW(250), "u") ' ú
    resultado = Replace(resultado, ChrW(249), "u") ' ù
    resultado = Replace(resultado, ChrW(251), "u") ' û
    resultado = Replace(resultado, ChrW(252), "u") ' ü

    resultado = Replace(resultado, ChrW(231), "c") ' ç

    RemoverAcentos = resultado
End Function



Private Sub apagarVestigios()

    Produtos.Range("AJ2", "Az1000").ClearContents


End Sub










