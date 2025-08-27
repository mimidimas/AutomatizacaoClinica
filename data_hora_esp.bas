Attribute VB_Name = "data_hora_esp"
Option Explicit
Sub especialidadem()
    Dim i As Integer 'contador
    Dim mensagem As String, g As Integer, ort As Integer, ot As Integer, d As Integer ' variavel para armazenar a especialidade
    
    i = 2
    g = 0
    ot = 0
    ort = 0
    d = 0
    
    Do While Cells(i, "B").Value <> Empty 'looping para verificar se a coluna Especialidade est� diferente de vazio
        Select Case Cells(i, "B").Value ' para fazer a verifica��o conforme a coluna B
        Case "Ortopedia" ' se estiver escrito Ortopedia
            ort = ort + 1
        Case "Ginencologia" ' se estiver escrito Ginencologia
           g = g + 1
        Case "Otorrinolaringologia" ' se estiver escrito Otorrinolaringologia
            ot = ot + 1
        Case "Dermatologia" ' se estiver escrito Dermatologia
            d = d + 1
        Case Else
             mensagem = "Cadastre os pacientes para poder ver a especialidade mais procurada"
        End Select 'finaliza o select case
        i = i + 1 ' add +1 no contador
    Loop ' temina o loop
    If g > ort And g > ot And g > d Then
        MsgBox "Ginencologia � a especialidade mais procurada"
    ElseIf ort > g And ort > ot And ort > d Then
        MsgBox "Ortopedia � a especialidade mais procurada"
    ElseIf ot > g And ot > ort And ot > d Then
        MsgBox "Ortopedia � a especialidade mais procurada"
    ElseIf d > g And d > ort And d > ot Then
        MsgBox "Dermatologia � a especialidade mais procurada"
    End If
    
    If ort + g + ot + d = 0 Then
    MsgBox mensagem, vbExclamation, "Aviso"
End If

    
End Sub

Sub horam()
    Dim i As Integer, m As Integer, t As Integer, n As Integer ' contador
    Dim turno As String ' receber o turno
    Dim hora As Date ' recebe a celula
    Dim mensagem As String ' recebe o turno
    
    i = 2
    Do While Cells(i, "D").Value <> Empty ' looping para verificar celulas com conte�do
        hora = Cells(i, "D").Value
        If hora >= TimeValue("07:00:00") And hora < TimeValue("12:00:00") Then   ' verificar se o hor�rio adicionado � no periodo da manh�
            m = m + 1 ' turno recebe manh�
        Else
            If hora >= TimeValue("12:00:00") And hora < TimeValue("18:00:00") Then ' verificar se o hor�rio adicionado � no periodo da tarde
                 t = t + 1 ' turno recebe Tarde
            Else
                If hora >= TimeValue("18:00:00") And hora < TimeValue("20:00:00") Then ' verificar se o hor�rio adicionado � no periodo da noite
                     n = n + 1 ' turno recebe Noite
                End If
            End If
        End If
        
        i = i + 1 'conta + 1 no contador
    Loop
    If m > t And m > n Then
             MsgBox "O turno mais procurado � o da manh�!"
        ElseIf t > m And t > n Then
            MsgBox "O turno mais procurado � o da tarde!"
        ElseIf n > m And n > t Then
            MsgBox "O turno mais procurado � o da noite!"
        End If
If Not IsDate(Cells(i, "D").Value) Then
    MsgBox "Hor�rio inv�lido encontrado na linha " & i & ".", vbExclamation, "Erro"
End If
End Sub

Function VerificarData(data As Date) As Boolean
    Dim dataAtual As Date
    dataAtual = Date ' Obt�m a data atual do sistema

    ' Verifica se a data fornecida � anterior � data atual
    If data < dataAtual Then
        VerificarData = False
    Else
        VerificarData = True
    End If
End Function

Sub pesquisar()
    Dim i As Integer
    Dim codigo_paciente As String
    
    ' Solicita o c�digo do paciente
    codigo_paciente = InputBox("Digite o c�digo do paciente:")
    
    ' Percorre as linhas para buscar o c�digo
    i = 2
    Do While Cells(i, "E").Value <> Empty
        If Cells(i, "E").Value = codigo_paciente Then
            MsgBox "Paciente encontrado:" & vbCrLf & "Nome: " & Cells(i, "A").Value & vbCrLf & "Especialidade: " & Cells(i, "B").Value & vbCrLf & "Data: " & Cells(i, "C").Text & vbCrLf & "Hora: " & Format(Cells(i, "D").Value, "hh:mm") & vbCrLf & "C�digo: " & Cells(i, "E").Value, vbInformation, "Informa��es do Paciente"
            Exit Sub
        End If
        i = i + 1
    Loop

    ' Caso n�o encontre, exibe mensagem
    MsgBox "Paciente n�o encontrado.", vbExclamation, "Aviso"
End Sub

