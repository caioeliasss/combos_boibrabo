Attribute VB_Name = "ModAtualizador"
Sub AtualizarVBA()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim URLModulo As String
    Dim CaminhoModulo As String
    Dim URLUserForm As String
    Dim CaminhoUserForm As String
    Dim URLFrx As String
    Dim CaminhoFrx As String
    Dim http As Object
    Dim fileNum As Integer
    Dim texto As String
    Dim modArray As Variant
    Dim userformArray As Variant
    Dim NomeModulo As String
    Dim NomeUserForm As String
    Dim i As Integer

    ' Array de módulos para atualizar
    modArray = Array( _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/modDatabase.bas", "modDatabase"), _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/modRange.bas", "modRange") _
    )
    
    ' Array de UserForms para atualizar (incluindo URLs para .frm e .frx)
    userformArray = Array( _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformAlterarAvulso.frm", "userformAlterarAvulso", "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformAlterarAvulso.frx"), _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformAlterarCombo.frm", "userformAlterarCombo", "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformAlterarCombo.frx"), _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformCriacaoCombos.frm", "userformCriacaoCombos", "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformCriacaoCombos.frx"), _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformVisualizacao.frm", "userformVisualizacao", "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/userformVisualizacao.frx"), _
        Array("https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/Calendario_V2.frm", "Calendario_V2", "https://raw.githubusercontent.com/caioeliasss/combos_boibrabo/main/Calendario_V2.frx"))
    
    ' -------------------------------
    ' Processar módulos (.bas)
    ' -------------------------------
    For i = LBound(modArray) To UBound(modArray)
        URLModulo = modArray(i)(0)
        NomeModulo = modArray(i)(1)
        CaminhoModulo = Environ("TEMP") & "\" & NomeModulo & ".bas"
        
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", URLModulo, False
        http.Send
        
        If http.status = 200 Then
            texto = http.responseText
        Else
            MsgBox "Erro ao baixar o módulo: " & NomeModulo, vbCritical
            Exit Sub
        End If
        
        fileNum = FreeFile
        Open CaminhoModulo For Output As #fileNum
        Print #fileNum, texto
        Close #fileNum
        
        Set vbProj = ThisWorkbook.VBProject
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(NomeModulo)
        On Error GoTo 0
        
        Set vbComp = vbProj.VBComponents.Import(CaminhoModulo)
        vbComp.Name = NomeModulo
    Next i
    
    ' -------------------------------
    ' Processar UserForms (.frm e .frx)
    ' -------------------------------
    For i = LBound(userformArray) To UBound(userformArray)
        URLUserForm = userformArray(i)(0)   ' URL do arquivo .frm
        NomeUserForm = userformArray(i)(1)    ' Nome do UserForm
        URLFrx = userformArray(i)(2)          ' URL do arquivo .frx
        
        CaminhoUserForm = Environ("TEMP") & "\" & NomeUserForm & ".frm"
        CaminhoFrx = Environ("TEMP") & "\" & NomeUserForm & ".frx"
        
        ' Baixar o arquivo .frm (como texto)
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", URLUserForm, False
        http.Send
        
        If http.status = 200 Then
            texto = http.responseText
        Else
            MsgBox "Erro ao baixar o UserForm (.frm): " & NomeUserForm, vbCritical
            Exit Sub
        End If
        
        fileNum = FreeFile
        Open CaminhoUserForm For Output As #fileNum
        Print #fileNum, texto
        Close #fileNum
        
        ' Baixar o arquivo .frx usando ADODB.Stream (para preservar os dados binários)
        If Not DownloadBinaryFile(URLFrx, CaminhoFrx) Then
            MsgBox "Erro ao baixar o arquivo .frx para o UserForm: " & NomeUserForm, vbCritical
            Exit Sub
        End If
        
        Set vbProj = ThisWorkbook.VBProject
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(NomeUserForm)
        On Error GoTo 0
        
        ' Importar o UserForm (o VBA automaticamente procurará o .frx correspondente no mesmo diretório)
        On Error Resume Next
        vbProj.VBComponents.Import (CaminhoUserForm)
        If Err.Number <> 0 Then
            MsgBox "Erro ao importar o UserForm: " & NomeUserForm & " - " & Err.Description, vbCritical
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    Next i
    
    ThisWorkbook.Save
    MsgBox "Atualização concluída com sucesso!", vbInformation
End Sub

' Função para baixar arquivos binários usando ADODB.Stream
Function DownloadBinaryFile(URL As String, FilePath As String) As Boolean
    Dim stm As Object
    Dim xmlHttp As Object
    On Error GoTo errHandler
    
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    xmlHttp.Open "GET", URL, False
    xmlHttp.Send
    
    If xmlHttp.status = 200 Then
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 1   ' 1 = adTypeBinary
        stm.Open
        stm.Write xmlHttp.responseBody
        stm.SaveToFile FilePath, 2   ' 2 = adSaveCreateOverWrite
        stm.Close
        Set stm = Nothing
        DownloadBinaryFile = True
    Else
        DownloadBinaryFile = False
    End If
    Set xmlHttp = Nothing
    Exit Function
    
errHandler:
    DownloadBinaryFile = False
End Function


