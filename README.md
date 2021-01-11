# projeto_excel_cadastro_clientes
Este é um projeto simples, que desenvolvi a um amigo para utilizar na empresa que atua. | This is a simple project, I devoloped for a friend to use in his job.

-- Código VBA --

Public Tecla As String


Private Sub CBEditar_Click()

On Error GoTo Erro

If TID = "" Or TextBox1 = "" Or TNome = "" Or TCpf = "" Or TTelefone = "" Or TEndereco = "" Then
MsgBox "Todos os campos precisam estar preenchidos!", vbCritical, "ERRO"
Exit Sub
End If

Dim ID As Double
ID = TID

Dim Data As Date
Data = TextBox1.Value

Dim Linha As Double
Linha = 1


With Plan1

    Do
    
    Linha = Linha + 1
    
    If .Cells(Linha, 2).Value = ID Then
        
        .Cells(Linha, 2).Value = ID
        .Cells(Linha, 3).Value = Data
        .Cells(Linha, 4).Value = TNome.Text
        .Cells(Linha, 5).Value = TCpf.Text
        .Cells(Linha, 6).Value = TTelefone.Text
        .Cells(Linha, 7).Value = TEndereco.Text

        Call Limpar

        MsgBox "Editado com Sucesso!", vbInformation, "EDITAR"
    
        Exit Sub
    
    End If
    
    Loop Until .Cells(Linha, 2).Value = ""
    
    MsgBox "Não encontrado!!", vbInformation, "EDITAR"
    
End With


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBExcluir_Click()

On Error GoTo Erro

If TID = "" Then
MsgBox "Necessário informar o ID!", vbCritical, "ERRO"
Exit Sub
End If

Dim Resposta As Integer

Resposta = MsgBox("Confirmar Exclusão?", VBA.vbYesNo, "EXCLUIR")

If Resposta = VBA.vbNo Then
Exit Sub
End If

Dim ID As Double
ID = TID


Dim Linha As Double
Linha = 1


With Plan1

    Do
    
    Linha = Linha + 1
    
    If .Cells(Linha, 2).Value = ID Then

        .Rows(Linha).Delete
    
        Call Limpar

        MsgBox "Excluído com Sucesso!", vbInformation, "EXCLUIR"
    
        Exit Sub
    
    End If
    
    Loop Until .Cells(Linha, 2).Value = ""
    
    MsgBox "Não encontrado!!", vbInformation, "EXCLUIR"
    
End With


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBLimpar_Click()

On Error GoTo Erro

Call Limpar

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBPesquisar_Click()

On Error GoTo Erro

Dim Pesquisa As String

Pesquisa = InputBox("Digite o critério de Pesquisa!", "PESQUISAR")

Dim Linha As Double
Linha = 1


With Plan1

    Do
    
    Linha = Linha + 1
    
    If .Cells(Linha, 2).Value = Pesquisa Or .Cells(Linha, 4).Value = Pesquisa _
    Or .Cells(Linha, 5).Value = Pesquisa Or .Cells(Linha, 6).Value = Pesquisa Then
    
            TID = .Cells(Linha, 2).Value
            TextBox1 = .Cells(Linha, 3).Value
            TNome.Text = .Cells(Linha, 4).Value
            TCpf.Text = .Cells(Linha, 5).Value
            TTelefone.Text = .Cells(Linha, 6).Value
            TEndereco.Text = .Cells(Linha, 7).Value
        
        Exit Sub
    
    End If
    
    Loop Until .Cells(Linha, 2).Value = ""
    
    MsgBox "Não encontrado!!", vbInformation, "PESQUISA"
    
End With

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBSalvar_Click()

On Error GoTo Erro

If TID = "" Then

TID = WorksheetFunction.Max(Plan1.Range("B1:B10000")) + 1
End If

Dim Ver As Double
Ver = WorksheetFunction.CountIf(Plan1.Range("B1:B10000"), TID)

If Ver > 0 Then
MsgBox "Códigos já cadastrado!", vbCritical, "ERRO"
Exit Sub
End If

If TData = "" Or TNome = "" Or TCpf = "" Or TTelefone = "" Or TEndereco = "" Then
MsgBox "Todos os campos precisam estar preenchidos!", vbCritical, "ERRO"
Exit Sub
End If

Dim Data As Date
Data = TextBox1.Value

Dim Linha As Double
Linha = 1

Dim ID As Double
ID = TID

With Plan1

    Do
    
    Linha = Linha + 1
    
    Loop Until .Cells(Linha, 2).Value = ""

.Cells(Linha, 2).Value = ID
.Cells(Linha, 3).Value = Data
.Cells(Linha, 4).Value = TNome.Text
.Cells(Linha, 5).Value = TCpf.Text
.Cells(Linha, 6).Value = TTelefone.Text
.Cells(Linha, 7).Value = TEndereco.Text

Call Limpar

MsgBox "Salvo com Sucesso!", vbInformation, "SALVAR"

End With



Exit Sub
Erro:
MsgBox "Erro ao Salvar!", vbCritical, "ERRO"

End Sub

Private Sub TCpf_Change()

On Error Resume Next
If Tecla = 8 Then
Exit Sub
End If

If VBA.Len(TCpf.Text) = 3 Then
TCpf = TCpf + "."
End If

If VBA.Len(TCpf.Text) = 7 Then
TCpf = TCpf + "."
End If

If VBA.Len(TCpf.Text) = 11 Then
TCpf = TCpf + "-"
End If


End Sub

Private Sub TData_Click()

If VBA.Len(TextBox1.Text) = 2 Then
TextBox1 = TextBox1 + "/"

If VBA.Len(TextBox1.Text) = 5 Then
TextBox1 = TextBox1 + "/"
End If

End Sub

Private Sub TEndereco_Change()

TEndereco = VBA.UCase(TEndereco.Text)

End Sub

Private Sub TextBox1_Change()

On Error Resume Next
If Tecla = 8 Then
Exit Sub
End If

If VBA.Len(TextBox1.Text) = 2 Then
TextBox1 = TextBox1 + "/"
End If

If VBA.Len(TextBox1.Text) = 5 Then
TextBox1 = TextBox1 + "/"
End If

End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Tecla = Empty
Tecla = KeyCode

End Sub

Private Sub TNome_Change()

TNome = VBA.UCase(TNome.Text)

End Sub

Private Sub TTelefone_Change()

On Error Resume Next
If Tecla = 8 Then
Exit Sub
End If

If VBA.Len(TTelefone.Text) = 1 Then
TTelefone = "(" + TTelefone
End If

If VBA.Len(TTelefone.Text) = 3 Then
TTelefone = TTelefone + ")"
End If

If VBA.Len(TTelefone.Text) = 9 Then
TTelefone = TTelefone + "-"
End If


End Sub

Private Sub UserForm_Initialize()

TextBox1 = VBA.Date


End Sub
Sub Limpar()

TID = ""
TData = ""
TNome = ""
TCpf = ""
TTelefone = ""
TEndereco = ""
TextBox1 = VBA.Date


End Sub

