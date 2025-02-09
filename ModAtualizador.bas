Sub AtualizarVBA()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim URLModulo As String
    Dim URLUserForm As String
    Dim CaminhoModulo As String
    Dim CaminhoUserForm As String
    Dim http As Object
    Dim fileNum As Integer
    Dim dados As String

    ' Definir URLs dos arquivos no Google Drive
    URLModulo = "https://drive.google.com/uc?export=download&id=ID_DO_MODULO"
    URLUserForm = "https://drive.google.com/uc?export=download&id=1E2RM00XGquIkkoEt322dzvxsDjuu6S4n"

    ' Definir caminhos temporários para salvar os arquivos baixados
    CaminhoModulo = Environ("TEMP") & "\Modulo1.bas"
    CaminhoUserForm = Environ("TEMP") & "\UserForm1.frm"

    ' Criar objeto para baixar o módulo
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URLModulo, False
    http.Send

    ' Se o download foi bem-sucedido, salvar o arquivo
    If http.Status = 200 Then
        fileNum = FreeFile
        Open CaminhoModulo For Output As #fileNum
        Print #fileNum, http.responseText
        Close #fileNum
    End If

    ' Criar objeto para baixar o UserForm
    http.Open "GET", URLUserForm, False
    http.Send

    ' Se o download foi bem-sucedido, salvar o arquivo
    If http.Status = 200 Then
        fileNum = FreeFile
        Open CaminhoUserForm For Output As #fileNum
        Print #fileNum, http.responseText
        Close #fileNum
    End If

    ' Atualizar os componentes VBA
    Set vbProj = ThisWorkbook.VBProject
    
    ' Remover módulo antigo e importar o novo
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("Modulo1")
    On Error GoTo 0
    vbProj.VBComponents.Import CaminhoModulo

    ' Remover UserForm antigo e importar o novo
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("UserForm1")
    On Error GoTo 0
    vbProj.VBComponents.Import CaminhoUserForm

    MsgBox "Código e UserForms atualizados com sucesso!", vbInformation
End Sub

