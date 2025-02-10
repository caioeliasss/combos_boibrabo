VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendario 
   Caption         =   "Calendario"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   OleObjectBlob   =   "Calendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub CommandButton1_Click()
    If CommandButton1.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton1.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
    Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton1.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton2_Click()
    If CommandButton2.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton2.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton2.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton3_Click()
    If CommandButton3.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton3.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton3.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton4_Click()
    If CommandButton4.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton4.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton4.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton5_Click()
    If CommandButton5.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton5.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton5.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton6_Click()
    If CommandButton6.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton6.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton6.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton7_Click()
    If CommandButton7.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton7.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton7.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton8_Click()
    If CommandButton8.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton8.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton8.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton9_Click()
    If CommandButton9.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton9.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton9.Caption))
    Call getDaysUpgraded
End Sub
    Private Sub CommandButton10_Click()
    If CommandButton10.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton10.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton10.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton11_Click()
    If CommandButton11.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton11.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton11.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton12_Click()
    If CommandButton12.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton12.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton12.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton13_Click()
    If CommandButton13.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton13.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton13.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton14_Click()
    If CommandButton14.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton14.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton14.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton15_Click()
    If CommandButton15.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton15.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton15.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton16_Click()
    If CommandButton16.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton16.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton16.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton17_Click()
    If CommandButton17.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton17.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton17.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton18_Click()
    If CommandButton18.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton18.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton18.Caption))
    Call getDaysUpgraded
End Sub
Private Sub CommandButton19_Click()
    If CommandButton19.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton19.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton19.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton20_Click()
    If CommandButton20.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton20.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton20.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton21_Click()
    If CommandButton21.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton21.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton21.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton22_Click()
    If CommandButton22.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton22.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton22.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton23_Click()
    If CommandButton23.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton23.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton23.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton24_Click()
    If CommandButton24.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton24.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton24.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton25_Click()
    If CommandButton25.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton25.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton25.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton26_Click()
    If CommandButton26.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton26.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton26.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton27_Click()
    If CommandButton27.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton27.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton27.Caption))
    Call getDaysUpgraded
End Sub
Private Sub CommandButton28_Click()
    If CommandButton28.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton28.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton28.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton29_Click()
    If CommandButton29.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton29.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton29.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton30_Click()
    If CommandButton30.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton30.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton30.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton31_Click()
    If CommandButton31.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton31.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton31.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton32_Click()
    If CommandButton32.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton32.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton32.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton33_Click()
    If CommandButton33.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton33.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton33.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton34_Click()
    If CommandButton34.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton34.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton34.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton35_Click()
    If CommandButton35.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton35.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If

    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton35.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton36_Click()
    If CommandButton36.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton36.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton36.Caption))
    Call getDaysUpgraded
End Sub


Private Sub CommandButton37_Click()
    If CommandButton37.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton37.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton37.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton38_Click()
    If CommandButton38.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton38.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton38.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton39_Click()
    If CommandButton39.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton39.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton39.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton40_Click()
    If CommandButton40.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton40.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton40.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton41_Click()
    If CommandButton41.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton41.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton41.Caption))
    Call getDaysUpgraded
End Sub

Private Sub CommandButton42_Click()
    If CommandButton42.BackColor = 11974327 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) - 1
        If mes = 0 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 1 Then ano = ano - 1
    ElseIf CommandButton42.BackColor = 11974326 Then
        mes = Val(Mid(labelDataAtual, 4, 2)) + 1
        If mes = 12 Then mes = 1
        ano = Val(Mid(labelDataAtual, 7, 4))
        If mes = 12 Then ano = ano + 1
    Else
       mes = Val(Mid(labelDataAtual, 4, 2))
       ano = Val(Mid(labelDataAtual, 7, 4))
    End If
Me.Hide
    labelDataSelecionada = DateSerial(ano, mes, Val(CommandButton42.Caption))
    Call getDaysUpgraded
End Sub


Private Sub buttonSelectRight_Click()
    Dim PrimeiroDiaMesSeguinte As Date
    PrimeiroDiaMesSeguinte = DateSerial(Year(labelDataAtual), month(labelDataAtual) + 1, 1)
    
    Call getHeaders(PrimeiroDiaMesSeguinte)
End Sub
Private Sub buttonSelectleft_Click()
    Dim PrimeiroDiaMesSeguinte As Date
    PrimeiroDiaMesSeguinte = DateSerial(Year(labelDataAtual), month(labelDataAtual) - 1, 1)
    
    Call getHeaders(PrimeiroDiaMesSeguinte)
    
    
End Sub

Private Sub Confirmar_Click()
Me.Hide
End Sub



Private Sub UserForm_Initialize()

Call getHeaders(Now)

End Sub

Sub getHeaders(dateSelected As Date)
Dim mes As String
Dim mesInt As Integer
Dim ano As Integer

labelDataSelecionada = ""

mesInt = month(dateSelected)
mes = getMesName(mesInt)
ano = Year(dateSelected)

labelDataAtual = DateSerial(Year(dateSelected), month(dateSelected), 1)
labelMes = mes
labelAno = ano

Call getDaysUpgraded

End Sub

Public Function getMesName(mes As Integer) As String

Select Case mes
    Case 1
        mesName = "Janeiro"
    Case 2
        mesName = "Fevereiro"
    Case 3
        mesName = "Março"
    Case 4
        mesName = "Abril"
    Case 5
        mesName = "Maio"
    Case 6
        mesName = "Junho"
    Case 7
        mesName = "Julho"
    Case 8
        mesName = "Agosto"
    Case 9
        mesName = "Setembro"
    Case 10
        mesName = "Outubro"
    Case 11
        mesName = "Novembro"
    Case 12
        mesName = "Dezembro"
End Select

getMesName = mesName

End Function

Private Sub getDaysUpgraded()
    Dim dataAtual As Date
    Dim i As Integer


    For i = 1 To 42
        Me.Controls("CommandButton" & i).BackColor = RGB(220, 220, 220)
    Next i
    
    ' Converte a string da label para uma data
    dataAtual = DateSerial(Val(Mid(labelDataAtual, 7, 4)), _
                           Val(Mid(labelDataAtual, 4, 2)), _
                           Val(Mid(labelDataAtual, 1, 2)))
    diaSemana = Weekday(dataAtual, vbSunday)
    countt = 0

    For i = diaSemana To 42
        Me.Controls("CommandButton" & i).Caption = Day(dataAtual + countt)
        If month(dataAtual + countt) <> month(dataAtual) Then
            Me.Controls("CommandButton" & i).BackColor = RGB(182, 182, 182)
        End If
        
        If dataAtual + countt = Date Then Me.Controls("CommandButton" & i).BackColor = RGB(173, 216, 230)
        If dataAtual + countt = labelDataSelecionada Then Me.Controls("CommandButton" & i).BackColor = RGB(255, 216, 230)
        If Weekday(dataAtual + countt) = 1 Then
            Me.Controls("CommandButton" & i).ForeColor = RGB(255, 0, 0)
            Me.Controls("CommandButton" & i).Font.Bold = True
        End If
        
        countt = countt + 1
        
        
        
    Next i
    
    countt = 0
    For i = 1 To diaSemana - 1
        Me.Controls("CommandButton" & i).Caption = Day(dataAtual - diaSemana + countt + 1)
        countt = countt + 1
        Me.Controls("CommandButton" & i).BackColor = RGB(183, 182, 182)
        If dataAtual = Now Then Me.Controls("CommandButton" & i).BackColor = RGB(173, 216, 230)
    Next i
    
    
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If labelDataSelecionada = "" Then
    MsgBox "Escolha uma data antes de sair."
    Cancel = 1
Else
    Me.Hide
    Cancel = 1
End If
End Sub


