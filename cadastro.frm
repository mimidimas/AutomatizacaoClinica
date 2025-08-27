VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastro 
   Caption         =   "cadastro"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   OleObjectBlob   =   "cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub continuar_Click()

    Dim i As Integer, planilha As Worksheet, j As Integer
    Dim dados As Variant
    Dim duplicado As Boolean

    ' Identificar a próxima linha vazia usando Do While
    i = 2 ' Começar na linha 2
    Do While Cells(i, "A").Value <> Empty
        i = i + 1
    Loop

    ' Verificar se o nome foi preenchido e se é válido
    If nometext.Value = "" Or TemNumero(nometext.Value) Then
        MsgBox "Digite um nome válido!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    ' Verificar se a data foi preenchida e é válida
    If text_data.Value = "" Or Not IsDate(text_data.Value) Then
        MsgBox "Digite uma data válida!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    ' Verificar se a data não é anterior à atual
    If Not VerificarData(CDate(text_data.Value)) Then
        MsgBox "A data não pode ser anterior à atual. Por favor, insira uma data válida!", vbCritical, "Data Inválida"
        Exit Sub
    End If

    ' Verificar se a hora foi preenchida e é válida
    If text_hora.Value = "" Or Not TemNumero(text_hora.Value) Then
        MsgBox "Digite uma hora válida!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    ' Verificar se a especialidade foi preenchida corretamente
    If cmd_esp.Value = "" Or (cmd_esp.Value <> "Ginecologia" And cmd_esp.Value <> "Otorrinolaringologia" And cmd_esp.Value <> "Ortopedia" And cmd_esp.Value <> "Dermatologia") Then
        MsgBox "Especialidade inválida! Selecione uma especialidade válida.", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    ' Verificar se o código do paciente foi preenchido corretamente
    If cdg_paciente.Value = "" Or Not TemNumero(cdg_paciente.Value) Then
        MsgBox "Digite um código de paciente válido!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

    ' Carregar os dados da planilha em uma matriz
    Set planilha = ThisWorkbook.Sheets("Cadastro")
    dados = planilha.Range("A2:E" & planilha.Cells(planilha.Rows.Count, "A").End(xlUp).Row).Value

    ' Verificar duplicidade
    duplicado = False
    For j = 1 To UBound(dados, 1)
        If dados(j, 1) = nometext.Value And _
           dados(j, 2) = cmd_esp.Value And _
           dados(j, 3) = text_data.Value And _
           dados(j, 4) = text_hora.Value And _
           dados(j, 5) = cdg_paciente.Value Then
            duplicado = True
            Exit For
        End If
    Next j

    If duplicado Then
        MsgBox "Cadastro já existente", vbCritical, "AVISO!"
        Exit Sub
    End If

    ' Inserir os dados no formulário
    Cells(i, "A").Value = nometext.Value
    Cells(i, "B").Value = cmd_esp.Value
    Cells(i, "C").Value = text_data.Value
    Cells(i, "D").Value = text_hora.Value
    Cells(i, "E").Value = cdg_paciente.Value

    MsgBox "Cadastro realizado com sucesso!", vbCritical

    ' Limpar os campos do UserForm
    nometext.Value = ""
    cmd_esp.Value = ""
    text_data.Value = ""
    text_hora.Value = ""
    cdg_paciente.Value = ""

End Sub


Sub LimparLinha(i As Integer)
         Cells(i, "A") = ""
         Cells(i, "B") = ""
         Cells(i, "C") = ""
         Cells(i, "D") = ""
         Cells(i, "E") = ""
End Sub

Private Sub CommandButton2_Click()

    Unload Me 'fecha o userform de cadastro

End Sub

Sub LimparClick_Click() 'limpa o userform

    nometext.Value = ""
    cmd_esp.Value = ""
    text_data.Value = ""
    text_hora.Value = ""

End Sub

Private Sub text_data_Change()
 Dim i As Integer
    i = Len(text_data.Text)  ' Conta quantos caracteres estão no TextBox
    
    ' Adiciona a barra após o segundo e o quinto caractere
    If i = 2 Or i = 5 Then
        text_data.Text = text_data.Text & "/"
    End If
    
    ' Coloca o cursor no final do texto
    text_data.SelStart = Len(text_data.Text)
End Sub

Private Sub text_hora_Change()
    Dim texto As String
    Dim partes() As String
    Dim horas As Integer, minutos As Integer
    
    ' Captura o texto atual
    texto = text_hora.Text
    
    ' Bloqueia caracteres não numéricos e os dois-pontos
    If Not texto Like "##:##" And texto <> "" Then
        If Len(texto) > 5 Or Not IsNumeric(Replace(texto, ":", "")) Then
            MsgBox "Digite apenas números no formato HH:MM.", vbCritical, "Hora Inválida"
            text_hora.Text = ""
            Exit Sub
        End If
    End If

    ' Adiciona o ':' automaticamente após o segundo caractere
    If Len(texto) = 2 And Not texto Like "##:" Then
        text_hora.Text = texto & ":"
        Exit Sub
    End If

    ' Validação de horas e minutos
    If Len(texto) = 5 Then
        partes = Split(texto, ":")
        
        If UBound(partes) = 1 Then
            horas = Val(partes(0))
            minutos = Val(partes(1))
            
            ' Verifica intervalo permitido para horas (07 a 20)
            If horas < 7 Or horas > 20 Then
                MsgBox "Digite uma hora entre 07:00 e 20:00.", vbCritical, "Hora Fora do Intervalo"
                text_hora.Text = ""
                Exit Sub
            End If
            
            ' Verifica intervalo permitido para minutos (00 a 59)
            If minutos < 0 Or minutos > 59 Then
                MsgBox "Digite minutos válidos (00 a 59).", vbCritical, "Minutos Inválidos"
                text_hora.Text = Left(texto, 3) ' Remove os minutos inválidos
                Exit Sub
            End If
        Else
            MsgBox "Formato inválido. Use o formato HH:MM.", vbCritical, "Formato Inválido"
            text_hora.Text = ""
            Exit Sub
        End If
    End If

    ' Coloca o cursor no final do texto
    text_hora.SelStart = Len(text_hora.Text)
End Sub

Private Sub UserForm_Initialize()
    cmd_esp.AddItem "Ginecologia" 'add os itens na caixa de combinção
    cmd_esp.AddItem "Otorrinolaringologia"
    cmd_esp.AddItem "Ortopedia"
    cmd_esp.AddItem "Dermatologia"
End Sub

Function TemNumero(str As String) As Boolean
    Dim i As Integer
    Dim caracter As String

    TemNumero = False

    For i = 1 To Len(str)
        caracter = Mid(str, i, 1)
        
        If IsNumeric(caracter) Then
            TemNumero = True
             Exit Function
        End If
    Next i
End Function

