VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Monitoramento - Tratamento de Água "
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBoxParametro_Change()
    MostrarOcultarFrame2
    Dim parametro As String
    parametro = ComboBoxParametro.Value
    
    ' Ocultar todos os controles
    LabelHorario.Visible = False
    TextBoxHorario.Visible = False
    LabelData.Visible = False
    TextBoxData.Visible = False
    LabelValor1.Visible = False
    TextBoxValor1.Visible = False
    LabelValor2.Visible = False
    TextBoxValor2.Visible = False
    LabelValor3.Visible = False
    TextBoxValor3.Visible = False
    LabelValor4.Visible = False
    TextBoxValor4.Visible = False
    
    ' Mostrar o campo de data para todos os parâmetros
    LabelData.Caption = "Data da medição:"
    LabelData.Visible = True
    TextBoxData.Visible = True
    
    ' Mostrar o campo de horário para todos os parâmetros
    LabelHorario.Caption = "Horário da medição:"
    LabelHorario.Visible = True
    TextBoxHorario.Visible = True
    
    ' Mostrar os controles relevantes com base no parâmetro selecionado
    Select Case parametro
        Case "pH"
            LabelValor1.Caption = "Osmose"
            LabelValor1.Visible = True
            TextBoxValor1.Visible = True
        Case "TOC"
            LabelValor1.Caption = "Retorno do loop"
            LabelValor1.Visible = True
            TextBoxValor1.Visible = True
        Case "Condutividade"
            LabelValor1.Caption = "Entrada UV-01 Entrada da Osmose"
            LabelValor1.Visible = True
            TextBoxValor1.Visible = True
            LabelValor2.Caption = "Saída da Osmose - 1º passo"
            LabelValor2.Visible = True
            TextBoxValor2.Visible = True
            LabelValor3.Caption = "Saída da Osmose - 2º passo"
            LabelValor3.Visible = True
            TextBoxValor3.Visible = True
            LabelValor4.Caption = "Saída para o loop"
            LabelValor4.Visible = True
            TextBoxValor4.Visible = True
        Case "Vazão"
            LabelValor1.Caption = "Entrada da Osmose - 1º passo"
            LabelValor1.Visible = True
            TextBoxValor1.Visible = True
            LabelValor2.Caption = "Saída da Osmose - 1º rejeito"
            LabelValor2.Visible = True
            TextBoxValor2.Visible = True
            LabelValor3.Caption = "Saída da Osmose - 2º rejeito"
            LabelValor3.Visible = True
            TextBoxValor3.Visible = True
            LabelValor4.Caption = "Produto"
            LabelValor4.Visible = True
            TextBoxValor4.Visible = True
    End Select
End Sub

Private Sub MostrarOcultarFrame2()
    If Not IsEmpty(ComboBoxSetor.Value) And Not IsEmpty(ComboBoxParametro.Value) Then
        Frame2.Visible = True
    Else
        Frame2.Visible = False
    End If
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub LabelValor3_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Inicialmente, o Frame2 deve estar oculto
    Frame2.Visible = False
    
    ComboBoxSetor.AddItem "STA 1"
    ComboBoxSetor.AddItem "STA 2"
    
    ComboBoxParametro.AddItem "pH"
    ComboBoxParametro.AddItem "Condutividade"
    ComboBoxParametro.AddItem "Vazão"
    ComboBoxParametro.AddItem "TOC"
End Sub

Private Sub ComboBoxParametro_DropButtonClick()
    If ComboBoxParametro.ListCount = 0 Then
        ComboBoxParametro.AddItem "pH"
        ComboBoxParametro.AddItem "Condutividade"
        ComboBoxParametro.AddItem "Vazão"
        ComboBoxParametro.AddItem "TOC"
    End If
End Sub
Private Function ControlExists(controlName As String) As Boolean
    On Error Resume Next
    ControlExists = Not Me.Controls(controlName) Is Nothing
    On Error GoTo 0
End Function

Private Sub CommandButtonOK_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim userName As String
    Dim setor As String
    Dim parametro As String
    Dim horario As String
    Dim data As String
    Dim i As Integer
    Dim camposIncompletos As Boolean

    Set ws = ThisWorkbook.Sheets("Dados")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    userName = Environ("USERNAME")
    setor = ComboBoxSetor.Value
    parametro = ComboBoxParametro.Value
    horario = TextBoxHorario.Text
    data = TextBoxData.Text

    camposIncompletos = False

    ' Verificar se todos os campos estão preenchidos
    If setor = "" Or parametro = "" Or horario = "" Or data = "" Then
        MsgBox "Por favor, preencha todos os campos!", vbExclamation
        Exit Sub
    End If

    ' Adicionar o número de linhas de acordo com a quantidade de LabelBox
    For i = 1 To 10 ' Assumindo que você tenha até 10 TextBoxValor, ajuste conforme necessário
        Dim textBoxName As String
        textBoxName = "TextBoxValor" & i
        If ControlExists(textBoxName) Then
            If Me.Controls(textBoxName).Text = "" Then
                camposIncompletos = True
            Else
                With ws
                    .Cells(lastRow, 1).Value = userName
                    .Cells(lastRow, 2).Value = setor
                    .Cells(lastRow, 3).Value = parametro
                    .Cells(lastRow, 4).Value = data
                    .Cells(lastRow, 5).Value = horario
                    .Cells(lastRow, 6).Value = Me.Controls("LabelValor" & i).Caption
                    .Cells(lastRow, 7).Value = Me.Controls(textBoxName).Text
                End With
                lastRow = lastRow + 1
            End If
        Else
            Exit For ' Sair do loop se não houver mais TextBoxValor
        End If
    Next i

    If camposIncompletos Then
        MsgBox "Por favor, preencha todos os campos!", vbExclamation
        Exit Sub
    End If

    MsgBox "Dados registrados com sucesso!", vbInformation

    ' Limpar os campos após o registro
    ComboBoxSetor.Value = ""
    ComboBoxParametro.Value = ""
    TextBoxHorario.Text = ""
    TextBoxData.Text = ""
    For i = 1 To 10
        textBoxName = "TextBoxValor" & i
        If ControlExists(textBoxName) Then
            Me.Controls(textBoxName).Text = ""
        Else
            Exit For ' Sair do loop se não houver mais TextBoxValor
        End If
    Next i
End Sub





Private Sub CommandButtonLimpar_Click()
    ' Limpar os TextBoxes
    TextBoxHorario.Text = ""
    TextBoxData.Text = ""
    TextBoxValor1.Text = ""
    TextBoxValor2.Text = ""
    TextBoxValor3.Text = ""
    TextBoxValor4.Text = ""
    
    ' Limpar os ComboBoxes
    ComboBoxSetor.Value = ""
    ComboBoxParametro.Value = ""
End Sub

