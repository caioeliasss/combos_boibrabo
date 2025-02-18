Attribute VB_Name = "modDatabase"
Public meuId As String
Public isValid As Boolean



Public Sub AtualizarDatabase()
    Dim var As Variant
    Dim var2 As Variant
    Dim caminhoArquivo As String

    
    
    caminhoArquivo = ThisWorkbook.Path & "\PRODUTOS.xlsx"
    
    Call createDocument(Log, Array("last_update", Now))
    
    Set wb = Workbooks.Open(caminhoArquivo)
    
    var = wb.Sheets(1).Range("A1").CurrentRegion.Value
    
    wb.Save
    wb.Close

    var2 = Produtos.Range("a1").CurrentRegion

    Produtos.Range("a1:l2000").ClearContents
    
    Produtos.Range("a1:l" & UBound(var)).Value = var
    count_ = 1
    For j = 2 To UBound(var)
        For i = 2 To UBound(var2)
            If var(j, 1) = var2(i, 1) Then
                Produtos.Cells(j, 13) = var2(i, 13)
                Produtos.Cells(j, 14) = var2(i, 14)
            End If
        Next i
    Next j
          

End Sub

Public Sub ConsultarPagamento()
    Dim http As Object
    Dim URL As String
    Dim dados As Variant
    Dim i As Integer
    Dim var As Variant
    Dim linha As Variant
    
    
    On Error Resume Next
    last_check = consultarDatabase(Log.Range("a1").CurrentRegion, Log, 1, "last_check", 2)
    On Error GoTo 0
    
    If last_check = "" Then
        isDue = 1000
    Else
        isDue = Now - last_check
    End If
    
    If isDue >= 5 Then
            
        
        URL = "https://docs.google.com/spreadsheets/d/1_MmDQ2Ei3xBqD-vyIp6icLEPCUz1IIaSpUbmN7OxfB0/export?format=csv"
        
        meuId = consultarDatabase(Log.Range("a1"), Log, 1, "meu_id", 2)
        
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", URL, False
        http.Send
        
        If http.status = 200 Then
            dados = Split(http.responseText, vbLf)
            
            ReDim var(0 To UBound(dados), 0 To 3)
            
            
            For i = 0 To UBound(dados)
                linha = Split(dados(i), ",")
                
                For col = 0 To UBound(linha)
                    var(i, col) = linha(col)
                Next col
            Next i
            
            
            primeiroDia = DateSerial(Year(Date), month(Date), 1)
            
            For i = 1 To UBound(var)
                If var(i, 0) = meuId And primeiroDia = CDate(var(i, 2)) Then
                    If var(i, 3) = "TRUE" Then
                        Call createDocument(Log, Array("last_check", Now))
                        isValid = True
                    Else
                        MsgBox ("Sua assinatura nao esta valida, voce nao tera mais acesso. Entre em contato com o distribuidor")
                        isValid = False
                    End If
                Else
                    MsgBox ("Sua assinatura nao esta valida, voce nao tera mais acesso. Entre em contato com o distribuidor")
                    isValid = False
                End If
            Next i
            
        Else
            MsgBox "Erro ao acessar os dados!", vbCritical
        End If
    Else
        isValid = True
    End If
    
End Sub

Public Sub createDocument(sheet As Worksheet, data As Variant)
    lr = sheet.UsedRange.Rows.Count + 1
    
    sheet.Range(Cells(lr, 1).Address, Cells(lr, UBound(data) + 1).Address).Value = data
    
End Sub

Private Function isArquivo(nome As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    log_database = ThisWorkbook.Path & "\" & nome & ".dat"
        
    
    If fso.fileExists(log_database) Then
        isArquivo = True
    Else
        isArquivo = False
    End If
    
End Function




