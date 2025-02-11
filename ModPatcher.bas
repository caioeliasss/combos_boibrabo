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

    ' Caminho tempor�rio para salvar o novo c�digo
    CaminhoModulo = Environ("TEMP") & "\ModAtualizador.bas"

    ' Criar objeto HTTP para baixar o novo c�digo
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URLModulo, False
    http.Send

    ' Se o download foi bem-sucedido, pegar o conte�do como texto
    If http.status = 200 Then
        texto = http.responseText
    Else
        MsgBox "Erro ao baixar o c�digo atualizado.", vbCritical
        Exit Sub
    End If

    ' Lendo o conte�do do arquivo para garantir a codifica��o correta
    ' Limpar quaisquer caracteres que possam ser invis�veis ou indesejados
    texto = Mid(texto, InStr(texto, "Sub ")) ' Ajuste para garantir que come�a na primeira Sub

    ' Salvar o arquivo limpo novamente
    fileNum = FreeFile
    Open CaminhoModulo For Output As #fileNum
    Print #fileNum, texto
    Close #fileNum

    ' Refer�ncia ao projeto VBA
    Set vbProj = ThisWorkbook.VBProject

    ' Remover o m�dulo antigo
    On Error Resume Next
    Set comb = vbProj.VBComponents("ModAtualizador1")
    erro = (comb Is Nothing) ' Se comb for Nothing, significa que o m�dulo n�o existe
    On Error GoTo 0
    
    ' Se erro = True, significa que "ModAtualizador1" n�o existe, ent�o usar "ModAtualizador2"
    If erro = True Then
        numero = 1
        Set comb = vbProj.VBComponents("ModAtualizador2")
    Else
        numero = 2
    End If
    
    On Error Resume Next
    'Do
        vbProj.VBComponents.Remove comb
        
        DoEvents
    'Loop Until comb Is Nothing
    On Error GoTo 0
    
    ' Importar o m�dulo atualizado
    Set vbComp = vbProj.VBComponents.Import(CaminhoModulo)

    ' Renomear o m�dulo importado
    vbComp.Name = "ModAtualizador" & numero ' Define o nome do m�dulo corretamente

    ' Fechar e reabrir o Excel automaticamente para aplicar as mudan�as
    ThisWorkbook.Save
    'Application.Quit
End Sub



