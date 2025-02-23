Attribute VB_Name = "ModPatcher"
Sub downloadPatcher()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim URLModulo As String
    Dim CaminhoModulo As String
    Dim http As Object
    Dim fileNum As Integer
    Dim texto As String
    Dim comp As Object
    ' URL do arquivo RAW no GitHub
    URLModulo = "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/ModAtualizador.bas"

    ' Caminho temporário para salvar o novo código
    CaminhoModulo = Environ("TEMP") & "\ModAtualizador.bas"

    ' Criar objeto HTTP para baixar o novo código
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URLModulo, False
    http.Send

    ' Se o download foi bem-sucedido, pegar o conteúdo como texto
    If http.status = 200 Then
        texto = http.responseText
    Else
        MsgBox "Erro ao baixar o código atualizado.", vbCritical
        Exit Sub
    End If

    ' Lendo o conteúdo do arquivo para garantir a codificação correta
    ' Limpar quaisquer caracteres que possam ser invisíveis ou indesejados
    texto = Mid(texto, InStr(texto, "Sub ")) ' Ajuste para garantir que começa na primeira Sub

    ' Salvar o arquivo limpo novamente
    fileNum = FreeFile
    Open CaminhoModulo For Output As #fileNum
    Print #fileNum, texto
    Close #fileNum

    ' Referência ao projeto VBA
    Set vbProj = ThisWorkbook.VBProject

    On Error Resume Next
    Set comb = vbProj.VBComponents("ModAtualizador1")
    erro = (comb Is Nothing)
    If erro = "" Then
        Set comb = vbProj.VBComponents("ModAtualizador2")
        erro = (comb Is Nothing)
        numero = 1
    Else
        numero = 2
    End If
    
    On Error GoTo 0
    
    
    
    On Error Resume Next
    'Do
        vbProj.VBComponents.Remove comb
        
        DoEvents
    'Loop Until comb Is Nothing
    On Error GoTo 0
    
    ' Importar o módulo atualizado
    Set vbComp = vbProj.VBComponents.Import(CaminhoModulo)

    ' Renomear o módulo importado
    vbComp.Name = "ModAtualizador" & numero ' Define o nome do módulo corretamente

    ' Fechar e reabrir o Excel automaticamente para aplicar as mudanças
    ThisWorkbook.Save
    'Application.Quit
End Sub



