Private Sub btnSalvar_Click()
    Dim ws As Worksheet, wsFicha As Worksheet
    Dim proximaLinha As Long, linFicha As Long, i As Long
    Dim totalPecas As Long
    Dim metrosTeoricos As Double, consumoReal As Double
    Dim metrosProjetados As Double, diferencaMetros As Double
    Dim camadas As Long
    Dim ehJaqueta As Boolean
    Dim modDigitado As String
    
    ' Variáveis das Quantidades Reais de Peças
    Dim qtdPP As Long, qtdP As Long, qtdM As Long, qtdG As Long
    Dim qtdGG As Long, qtdEX As Long, qtdEXG As Long, qtd2G As Long
    Dim qtd3G As Long, qtd4G As Long, qtd5G As Long, qtdUnico As Long
    
    ' Define as abas de trabalho
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LANCAMENTOS")
    Set wsFicha = ThisWorkbook.Sheets("FICHA_TECNICA")
    On Error GoTo 0
    
    If ws Is Nothing Or wsFicha Is Nothing Then
        MsgBox "Atenção: Verifique se as abas 'LANCAMENTOS' e 'FICHA_TECNICA' existem na planilha!", vbCritical, "Erro de Configuração"
        Exit Sub
    End If
    
    proximaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Validação da Data
    If Not IsDate(Me.txtData.Value) Then
        MsgBox "Por favor, digite uma data válida (ex: 04/05/2026).", vbExclamation, "Data Inválida"
        Me.txtData.SetFocus
        Exit Sub
    End If
    
    modDigitado = Trim(Me.txtModelo.Value)
    camadas = CLng(IIf(Me.txtCamadas.Value = "", 0, Me.txtCamadas.Value))
    
    ' LÓGICA DAS JAQUETAS (Modelos terminados em 2 ou o modelo 2903)
    If Right(modDigitado, 1) = "2" Or modDigitado = "2903" Then
        ehJaqueta = True
    Else
        ehJaqueta = False
    End If
    
    If ehJaqueta And camadas > 0 Then
        qtdPP = Val(Me.txtPP.Value) * camadas
        qtdP = Val(Me.txtP.Value) * camadas
        qtdM = Val(Me.txtM.Value) * camadas
        qtdG = Val(Me.txtG.Value) * camadas
        qtdGG = Val(Me.txtGG.Value) * camadas
        qtdEX = Val(Me.txtEX.Value) * camadas
        qtdEXG = Val(Me.txtEXG.Value) * camadas
        qtd2G = Val(Me.txt2G.Value) * camadas
        qtd3G = Val(Me.txt3G.Value) * camadas
        qtd4G = Val(Me.txt4G.Value) * camadas
        qtd5G = Val(Me.txt5G.Value) * camadas
        qtdUnico = Val(Me.txtUnico.Value) * camadas
    Else
        qtdPP = Val(Me.txtPP.Value)
        qtdP = Val(Me.txtP.Value)
        qtdM = Val(Me.txtM.Value)
        qtdG = Val(Me.txtG.Value)
        qtdGG = Val(Me.txtGG.Value)
        qtdEX = Val(Me.txtEX.Value)
        qtdEXG = Val(Me.txtEXG.Value)
        qtd2G = Val(Me.txt2G.Value)
        qtd3G = Val(Me.txt3G.Value)
        qtd4G = Val(Me.txt4G.Value)
        qtd5G = Val(Me.txt5G.Value)
        qtdUnico = Val(Me.txtUnico.Value)
    End If
    
    totalPecas = qtdPP + qtdP + qtdM + qtdG + qtdGG + qtdEX + qtdEXG + qtd2G + qtd3G + qtd4G + qtd5G + qtdUnico
    metrosTeoricos = CDbl(IIf(Me.txtMedida.Value = "", 0, Replace(Me.txtMedida.Value, ".", ","))) * camadas
    
    If totalPecas > 0 Then
        consumoReal = metrosTeoricos / totalPecas
    Else
        consumoReal = 0
    End If
    
    ' BUSCA INTELIGENTE NA FICHA TÉCNICA (Compara Texto e Número)
    linFicha = 0
    For i = 2 To wsFicha.Cells(wsFicha.Rows.Count, "A").End(xlUp).Row
        If Trim(CStr(wsFicha.Cells(i, 1).Value)) = modDigitado Then
            linFicha = i
            Exit For
        End If
    Next i
    
    ' CÁLCULO DOS METROS PROJETADOS
    metrosProjetados = 0
    If linFicha > 0 Then
        metrosProjetados = (qtdPP * CDbl(IIf(wsFicha.Cells(linFicha, 2).Value = "", 0, wsFicha.Cells(linFicha, 2).Value))) + _
                           (qtdP * CDbl(IIf(wsFicha.Cells(linFicha, 3).Value = "", 0, wsFicha.Cells(linFicha, 3).Value))) + _
                           (qtdM * CDbl(IIf(wsFicha.Cells(linFicha, 4).Value = "", 0, wsFicha.Cells(linFicha, 4).Value))) + _
                           (qtdG * CDbl(IIf(wsFicha.Cells(linFicha, 5).Value = "", 0, wsFicha.Cells(linFicha, 5).Value))) + _
                           (qtdGG * CDbl(IIf(wsFicha.Cells(linFicha, 6).Value = "", 0, wsFicha.Cells(linFicha, 6).Value))) + _
                           (qtdEX * CDbl(IIf(wsFicha.Cells(linFicha, 7).Value = "", 0, wsFicha.Cells(linFicha, 7).Value))) + _
                           (qtdEXG * CDbl(IIf(wsFicha.Cells(linFicha, 8).Value = "", 0, wsFicha.Cells(linFicha, 8).Value))) + _
                           (qtd2G * CDbl(IIf(wsFicha.Cells(linFicha, 9).Value = "", 0, wsFicha.Cells(linFicha, 9).Value))) + _
                           (qtd3G * CDbl(IIf(wsFicha.Cells(linFicha, 10).Value = "", 0, wsFicha.Cells(linFicha, 10).Value))) + _
                           (qtd4G * CDbl(IIf(wsFicha.Cells(linFicha, 11).Value = "", 0, wsFicha.Cells(linFicha, 11).Value))) + _
                           (qtd5G * CDbl(IIf(wsFicha.Cells(linFicha, 12).Value = "", 0, wsFicha.Cells(linFicha, 12).Value))) + _
                           (qtdUnico * CDbl(IIf(wsFicha.Cells(linFicha, 13).Value = "", 0, wsFicha.Cells(linFicha, 13).Value)))
    Else
        MsgBox "Aviso: O modelo '" & modDigitado & "' não foi encontrado na aba FICHA_TECNICA.", vbWarning, "Modelo Não Cadastrado"
    End If
    
    diferencaMetros = metrosTeoricos - metrosProjetados
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    ' GRAVAÇÃO DOS DADOS
    ws.Cells(proximaLinha, 1).Value = CDate(Me.txtData.Value)
    ws.Cells(proximaLinha, 2).Value = modDigitado
    ws.Cells(proximaLinha, 3).Value = CDbl(IIf(Me.txtMedida.Value = "", 0, Replace(Me.txtMedida.Value, ".", ",")))
    ws.Cells(proximaLinha, 4).Value = camadas
    
    ws.Cells(proximaLinha, 5).Value = qtdPP
    ws.Cells(proximaLinha, 6).Value = qtdP
    ws.Cells(proximaLinha, 7).Value = qtdM
    ws.Cells(proximaLinha, 8).Value = qtdG
    ws.Cells(proximaLinha, 9).Value = qtdGG
    ws.Cells(proximaLinha, 10).Value = qtdEX
    ws.Cells(proximaLinha, 11).Value = qtdEXG
    ws.Cells(proximaLinha, 12).Value = qtd2G
    ws.Cells(proximaLinha, 13).Value = qtd3G
    ws.Cells(proximaLinha, 14).Value = qtd4G
    ws.Cells(proximaLinha, 15).Value = qtd5G
    ws.Cells(proximaLinha, 16).Value = qtdUnico
    
    ws.Cells(proximaLinha, 17).Value = totalPecas
    ws.Cells(proximaLinha, 18).Value = metrosTeoricos
    ws.Cells(proximaLinha, 19).Value = consumoReal
    ws.Cells(proximaLinha, 20).Value = metrosProjetados
    ws.Cells(proximaLinha, 21).Value = diferencaMetros
    
    On Error Resume Next
    ws.Protect AllowFiltering:=True, AllowSorting:=True
    On Error GoTo 0
    
    MsgBox "Lançamento do Modelo " & modDigitado & " salvo com sucesso!", vbInformation, "Sucesso"
    
    Me.txtModelo.Value = ""
    Me.txtMedida.Value = ""
    Me.txtCamadas.Value = ""
    Me.txtPP.Value = "": Me.txtP.Value = "": Me.txtM.Value = "": Me.txtG.Value = ""
    Me.txtGG.Value = "": Me.txtEX.Value = "": Me.txtEXG.Value = "": Me.txt2G.Value = ""
    Me.txt3G.Value = "": Me.txt4G.Value = "": Me.txt5G.Value = "": Me.txtUnico.Value = ""
    
    Me.txtModelo.SetFocus
End Sub

