Attribute VB_Name = "Routing"
Sub FillCities()
    Dim numberOfCities As Variant
    Dim i As Integer
    Dim lastRow As Long

    ' Determine the last filled row in column B
    lastRow = Cells(Rows.Count, 2).End(xlUp).row
    
    ' Request the number of cities from the user
    Do
        numberOfCities = Application.InputBox("Enter the number of new cities to be served:", Type:=1)
        
        ' Check if the user canceled the input
        If numberOfCities = False Then Exit Sub
        
        ' Check if the number of cities is valid
        If numberOfCities <= 0 Then
            MsgBox "The number of cities must be greater than zero."
        End If
    Loop While numberOfCities <= 0
    
    ' Fill in the city names in column B
    For i = 1 To numberOfCities
        Dim cityName As Variant
        Dim cityCode As String
        
        Do
            cityName = Application.InputBox("Enter the name of city " & i & ":")
            
            ' Check if the user canceled the input
            If cityName = False Then Exit Sub
            
            ' Check if the city name is valid
            If cityName = "" Then
                MsgBox "The city name cannot be empty."
            End If
            
            ' Check if the city has already been entered
            If CheckExistingCity(cityName, i) Then
                MsgBox "The city '" & cityName & "' has already been entered. Please enter another city."
                cityName = ""
            End If
        Loop While cityName = ""
        
        Cells(lastRow + i, 2).value = cityName
        cityCode = EncodeCity(lastRow + i - 2)
        Cells(lastRow + i, 3).value = cityCode
    Next i
End Sub

Function CheckExistingCity(ByVal cityName As String, ByVal cityNumber As Integer) As Boolean
    Dim rng As Range
    Dim cell As Range
    
    ' Define the range of cities already entered
    Set rng = Range("B3:B" & cityNumber + 2)
    
    ' Check if the city already exists in the range
    For Each cell In rng
        If UCase(cell.value) = UCase(cityName) Then
            CheckExistingCity = True
            Exit Function
        End If
    Next cell
    
    CheckExistingCity = False
End Function

Function EncodeCity(ByVal cityNumber As Integer) As String
    Dim modulo, division As Integer
    Dim code As String
    
    ' Determine the modulo and division to encode the city
    modulo = (cityNumber - 1) Mod 26 + 1
    division = (cityNumber - 1) \ 26
    
    ' Convert the modulo to the corresponding letter
    code = Chr(64 + modulo)
    
    ' Add additional letters if necessary
    Do While division > 0
        modulo = (division - 1) Mod 26
        code = Chr(64 + modulo + 1) & code
        division = (division - 1) \ 26
    Loop
    
    EncodeCity = code
End Function

Sub FillFuels()
    Dim numberOfFuels As Integer
    Dim i As Integer
    Dim fuel As String
    Dim fuelValue As Variant
    Dim lastRow As Long

    ' Determine the last filled row in column N
    lastRow = Cells(Rows.Count, 14).End(xlUp).row
    
    ' List of available fuels
    Dim availableFuels As Variant
    availableFuels = Array("", "Types of fuels:", "Regular gasoline", "Additive gasoline", "Formulated gasoline", "Ethanol", "Additive ethanol", "CNG", "Diesel S-500", "Diesel S-10", "Additive diesel", "Premium diesel")
    
    ' Request the number of fuels from the user
    Do
        numberOfFuels = Application.InputBox("Enter the number of fuels:", Type:=1)
        
        ' Check if the user canceled the input
        If numberOfFuels = False Then Exit Sub
        
        ' Check if the number of fuels is valid
        If numberOfFuels <= 0 Then
            MsgBox "The number of fuels must be greater than zero."
        End If
    Loop While numberOfFuels <= 0
    
    ' Fill in the names of the fuels in column N
    For i = 1 To numberOfFuels
        ' Show the list of available fuels
        fuel = Application.InputBox("Enter the name of fuel " & i & ":" & vbCrLf & Join(availableFuels, vbCrLf), Type:=2)
        
        ' Check if the user canceled the input
        If fuel = False Then Exit Sub
        
        ' Check if the entered fuel is valid
        If Not IsInArray(fuel, availableFuels) Then
            MsgBox "Invalid fuel. Please enter again."
            i = i - 1 ' Reduce the counter to repeat the iteration
        Else
            ' Fill in the name of the fuel in column N
            Cells(lastRow + i, 14).value = fuel
            
            ' Request the value of the fuel
            fuelValue = Application.InputBox("Enter the value of fuel " & fuel & ":", Type:=1)
            
            ' Check if the user canceled the input
            If fuelValue = False Then Exit Sub
            
            ' Fill in the value of the fuel in column O
            Cells(lastRow + i, 15).value = CDbl(fuelValue)
        End If
    Next i
End Sub

Function IsInArray(ByVal value As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(element, value, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
End Function

Sub CopyCityData()
    Dim wsRegistration As Worksheet
    Dim wsCurrent As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim transposedData As Variant
    Dim i As Long
    
    ' Define the source (Registration) and destination (current sheet) worksheets
    Set wsRegistration = ThisWorkbook.Worksheets("Registration")
    Set wsCurrent = ThisWorkbook.ActiveSheet
    
    ' Clear the data and formats in column D starting from D3
    wsCurrent.Range("D3:D" & wsCurrent.Cells(wsCurrent.Rows.Count, "D").End(xlUp).row).ClearContents
    wsCurrent.Range("D3:D" & wsCurrent.Cells(wsCurrent.Rows.Count, "D").End(xlUp).row).Borders(xlEdgeRight).LineStyle = xlNone
    
    ' Clear the formatting in row 2 starting from E2
    wsCurrent.Range("E2:Z2").Borders(xlEdgeBottom).LineStyle = xlNone
    
    ' Determine the last row with data in column C of the Registration sheet
    lastRow = wsRegistration.Cells(wsRegistration.Rows.Count, "C").End(xlUp).row
    
    ' Define the source range (column C of the Registration sheet)
    Set sourceRange = wsRegistration.Range("C3:C" & lastRow)
    
    ' Define the destination range in column D (starting from D3) in the current sheet
    Set destinationRange = wsCurrent.Range("D3").Resize(sourceRange.Rows.Count)
    
    ' Copy the data from column C of the Registration sheet to column D in the current sheet
    sourceRange.Copy Destination:=destinationRange
    
    ' Store the copied data in an array and transpose the data
    transposedData = Application.Transpose(destinationRange.value)
    
    ' Paste the transposed data in row 2 starting from column E
    For i = 1 To UBound(transposedData)
        wsCurrent.Cells(2, i + 4).value = transposedData(i)
    Next i
    
    ' Apply thick border formatting to the filled rows in column D
    With wsCurrent.Range("D3:D" & wsCurrent.Cells(wsCurrent.Rows.Count, "D").End(xlUp).row)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThick
    End With
    
    ' Apply thick bottom border formatting below the filled data in row 2
    With wsCurrent.Range("E2").Resize(1, UBound(transposedData))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
End Sub


Sub FillDemands()
    Dim numberOfDemands As Integer
    Dim cityName As String
    Dim demand As Double
    Dim i As Integer
    
    ' Request the number of demands from the user
    Do
        numberOfDemands = Val(InputBox("Enter the number of demands to be registered:"))
        
        ' Check if the user canceled the input
        If numberOfDemands = 0 Then Exit Sub
        
        ' Check if the number of demands is valid
        If numberOfDemands <= 0 Then
            MsgBox "The number of demands must be greater than zero."
        End If
    Loop While numberOfDemands <= 0
    
    ' Fill in the demands
    i = 1 ' Variable to control the count of valid demands
    
    Do While i <= numberOfDemands
        ' Request the name of the city from the user
        cityName = InputBox("Enter the name of city " & i & ":")
        
        ' Check if the user canceled the input
        If cityName = "" Then Exit Do
        
        ' Search for the city name in column B
        Dim cityCell As Range
        Set cityCell = Range("B:B").Find(What:=cityName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Check if the city was found
        If cityCell Is Nothing Then
            MsgBox "The city '" & cityName & "' was not found in the list."
        Else
            ' Request the demand from the user
            Do
                demand = Val(InputBox("Enter the demand in kg for the city " & cityName & ":"))
                
                ' Check if the user canceled the input
                If demand = 0 Then Exit Do
                
                ' Check if the demand is valid
                If demand < 0 Then
                    MsgBox "The demand must be a value equal to or greater than zero."
                End If
            Loop While demand < 0
            
            ' Fill in the demand in column D
            cityCell.Offset(0, 2).value = demand
            
            i = i + 1 ' Increment the count of valid demands
        End If
    Loop
End Sub

Sub FormatColumns()
'
' FormatColumns Macro
'

'
    ' Select columns B to O
    Columns("B:O").Select
    
    ' Activate cell O1 (the top of the selected range)
    Range("O1").Activate
    
    ' Autofit the width of columns B to O based on their content
    Columns("B:O").EntireColumn.AutoFit
    
    ' Select cell B2 to return to a specific starting point
    Range("B2").Select
End Sub

Sub PreencherCidades()
    Dim quantidadeCidades As Variant
    Dim i As Integer
    Dim lastRow As Long

    ' Determinar a última linha preenchida na coluna B
    lastRow = Cells(Rows.Count, 2).End(xlUp).row
    
    ' Solicitar a quantidade de cidades ao usuário
    Do
        quantidadeCidades = Application.InputBox("Digite a quantidade de novas cidades a serem atendidas:", Type:=1)
        
        ' Verificar se o usuário cancelou a entrada
        If quantidadeCidades = False Then Exit Sub
        
        ' Verificar se a quantidade de cidades é válida
        If quantidadeCidades <= 0 Then
            MsgBox "A quantidade de cidades deve ser maior que zero."
        End If
    Loop While quantidadeCidades <= 0
    
    ' Preencher os nomes das cidades na coluna B
    For i = 1 To quantidadeCidades
        Dim nomeCidade As Variant
        Dim codCidade As String
        
        Do
            nomeCidade = Application.InputBox("Digite o nome da cidade " & i & ":")
            
            ' Verificar se o usuário cancelou a entrada
            If nomeCidade = False Then Exit Sub
            
            ' Verificar se o nome da cidade é válido
            If nomeCidade = "" Then
                MsgBox "O nome da cidade não pode estar vazio."
            End If
            
            ' Verificar se a cidade já foi digitada
            If VerificarCidadeExistente(nomeCidade, i) Then
                MsgBox "A cidade '" & nomeCidade & "' já foi digitada. Digite outra cidade."
                nomeCidade = ""
            End If
        Loop While nomeCidade = ""
        
        Cells(lastRow + i, 2).value = nomeCidade
        codCidade = CodificarCidade(lastRow + i - 2)
        Cells(lastRow + i, 3).value = codCidade
    Next i
End Sub

Function VerificarCidadeExistente(ByVal nomeCidade As String, ByVal numCidade As Integer) As Boolean
    Dim rng As Range
    Dim cell As Range
    
    ' Definir o intervalo das cidades já digitadas
    Set rng = Range("B3:B" & numCidade + 2)
    
    ' Verificar se a cidade já existe no intervalo
    For Each cell In rng
        If UCase(cell.value) = UCase(nomeCidade) Then
            VerificarCidadeExistente = True
            Exit Function
        End If
    Next cell
    
    VerificarCidadeExistente = False
End Function

Function CodificarCidade(ByVal numCidade As Integer) As String
    Dim modulo, divisao As Integer
    Dim cod As String
    
    ' Determinar o módulo e a divisão para codificar a cidade
    modulo = (numCidade - 1) Mod 26 + 1
    divisao = (numCidade - 1) \ 26
    
    ' Converter o módulo para a letra correspondente
    cod = Chr(64 + modulo)
    
    ' Adicionar as letras adicionais, se necessário
    Do While divisao > 0
        modulo = (divisao - 1) Mod 26
        cod = Chr(64 + modulo + 1) & cod
        divisao = (divisao - 1) \ 26
    Loop
    
    CodificarCidade = cod
End Function

Sub PreencherCombustiveis()
    Dim quantidadeCombustiveis As Integer
    Dim i As Integer
    Dim combustivel As String
    Dim valorCombustivel As Variant
    Dim lastRow As Long

    ' Determinar a última linha preenchida na coluna N
    lastRow = Cells(Rows.Count, 14).End(xlUp).row
    
    ' Lista de combustíveis disponíveis
    Dim combustiveisDisponiveis As Variant
    combustiveisDisponiveis = Array("", "Tipos de combustíveis:", "Gasolina comum", "Gasolina aditivada", "Gasolina formulada", "Etanol", "Etanol aditivado", "GNV", "Diesel S-500", "Diesel S-10", "Diesel aditivada", "Diesel premium")
    
    ' Solicitar a quantidade de combustíveis ao usuário
    Do
        quantidadeCombustiveis = Application.InputBox("Digite a quantidade de combustíveis:", Type:=1)
        
        ' Verificar se o usuário cancelou a entrada
        If quantidadeCombustiveis = False Then Exit Sub
        
        ' Verificar se a quantidade de combustíveis é válida
        If quantidadeCombustiveis <= 0 Then
            MsgBox "A quantidade de combustíveis deve ser maior que zero."
        End If
    Loop While quantidadeCombustiveis <= 0
    
    ' Preencher os nomes dos combustíveis na coluna N
    For i = 1 To quantidadeCombustiveis
        ' Exibir a lista de combustíveis disponíveis
        combustivel = Application.InputBox("Digite o nome do combustível " & i & ":" & vbCrLf & Join(combustiveisDisponiveis, vbCrLf), Type:=2)
        
        ' Verificar se o usuário cancelou a entrada
        If combustivel = False Then Exit Sub
        
        ' Verificar se o combustível digitado é válido
        If Not IsInArray(combustivel, combustiveisDisponiveis) Then
            MsgBox "Combustível inválido. Digite novamente."
            i = i - 1 ' Reduzir o contador para repetir a iteração
        Else
            ' Preencher o nome do combustível na coluna N
            Cells(lastRow + i, 14).value = combustivel
            
            ' Solicitar o valor do combustível
            valorCombustivel = Application.InputBox("Digite o valor do combustível " & combustivel & ":", Type:=1)
            
            ' Verificar se o usuário cancelou a entrada
            If valorCombustivel = False Then Exit Sub
            
            ' Preencher o valor do combustível na coluna O
            Cells(lastRow + i, 15).value = CDbl(valorCombustivel)
        End If
    Next i
End Sub

Function IsInArray(ByVal value As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(element, value, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
End Function

Sub CopiarDadosCidades()
    Dim wsCadastro As Worksheet
    Dim wsAtual As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim transposedData As Variant
    Dim i As Long
    
    ' Definir as planilhas de origem (Cadastro) e destino (planilha atual)
    Set wsCadastro = ThisWorkbook.Worksheets("Cadastro")
    Set wsAtual = ThisWorkbook.ActiveSheet
    
    ' Limpar os dados e as formatações da coluna D a partir de D3
    wsAtual.Range("D3:D" & wsAtual.Cells(wsAtual.Rows.Count, "D").End(xlUp).row).ClearContents
    wsAtual.Range("D3:D" & wsAtual.Cells(wsAtual.Rows.Count, "D").End(xlUp).row).Borders(xlEdgeRight).LineStyle = xlNone
    
    ' Limpar as formatações da linha 2 a partir de E2
    wsAtual.Range("E2:Z2").Borders(xlEdgeBottom).LineStyle = xlNone
    
    ' Determinar a última linha com dados na coluna C da planilha Cadastro
    lastRow = wsCadastro.Cells(wsCadastro.Rows.Count, "C").End(xlUp).row
    
    ' Definir o intervalo de origem (coluna C da planilha Cadastro)
    Set sourceRange = wsCadastro.Range("C3:C" & lastRow)
    
    ' Definir o intervalo de destino na coluna D (a partir de D3) na planilha atual
    Set destinationRange = wsAtual.Range("D3").Resize(sourceRange.Rows.Count)
    
    ' Copiar os dados da coluna C da planilha Cadastro para a coluna D na planilha atual
    sourceRange.Copy Destination:=destinationRange
    
    ' Armazenar os dados copiados em uma matriz e transpor os dados
    transposedData = Application.Transpose(destinationRange.value)
    
    ' Colar os dados transpostos na linha 2 a partir da coluna E
    For i = 1 To UBound(transposedData)
        wsAtual.Cells(2, i + 4).value = transposedData(i)
    Next i
    
    ' Aplicar formatação de bordas espessas nas linhas preenchidas da coluna D
    With wsAtual.Range("D3:D" & wsAtual.Cells(wsAtual.Rows.Count, "D").End(xlUp).row)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThick
    End With
    
    ' Aplicar formatação de bordas espessas abaixo dos dados preenchidos na linha 2
    With wsAtual.Range("E2").Resize(1, UBound(transposedData))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
End Sub

Sub PreencherDemandas()
    Dim quantidadeDemandas As Integer
    Dim nomeCidade As String
    Dim demanda As Double
    Dim i As Integer
    
    ' Solicitar a quantidade de demandas ao usuário
    Do
        quantidadeDemandas = Val(InputBox("Digite a quantidade de demandas a serem cadastradas:"))
        
        ' Verificar se o usuário cancelou a entrada
        If quantidadeDemandas = 0 Then Exit Sub
        
        ' Verificar se a quantidade de demandas é válida
        If quantidadeDemandas <= 0 Then
            MsgBox "A quantidade de demandas deve ser maior que zero."
        End If
    Loop While quantidadeDemandas <= 0
    
    ' Preencher as demandas
    i = 1 ' Variável para controle da contagem de demandas válidas
    
    Do While i <= quantidadeDemandas
        ' Solicitar o nome da cidade ao usuário
        nomeCidade = InputBox("Digite o nome da cidade " & i & ":")
        
        ' Verificar se o usuário cancelou a entrada
        If nomeCidade = "" Then Exit Do
        
        ' Procurar o nome da cidade na coluna B
        Dim cidadeCelula As Range
        Set cidadeCelula = Range("B:B").Find(What:=nomeCidade, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verificar se a cidade foi encontrada
        If cidadeCelula Is Nothing Then
            MsgBox "A cidade '" & nomeCidade & "' não foi encontrada na lista."
        Else
            ' Solicitar a demanda ao usuário
            Do
                demanda = Val(InputBox("Digite a demanda em kg para a cidade " & nomeCidade & ":"))
                
                ' Verificar se o usuário cancelou a entrada
                If demanda = 0 Then Exit Do
                
                ' Verificar se a demanda é válida
                If demanda < 0 Then
                    MsgBox "A demanda deve ser um valor igual ou maior que zero."
                End If
            Loop While demanda < 0
            
            ' Preencher a demanda na coluna D
            cidadeCelula.Offset(0, 2).value = demanda
            
            i = i + 1 ' Incrementar a contagem de demandas válidas
        End If
    Loop
End Sub

Sub PreencherEntregas()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Entregas")
    
    ' Pede ao usuário a quantidade de entregas
    Dim qtdeEntregas As Integer
    qtdeEntregas = InputBox("Digite a quantidade de entregas a serem cadastradas:")
    
    Dim i As Integer
    Dim linhaDestino As Integer
    linhaDestino = 3 ' Linha inicial para colar os dados
    
    For i = 1 To qtdeEntregas
        ' Pede ao usuário o CEP da entrega
        Dim cep As String
        Dim rng As Range
        Do
            cep = InputBox("Digite o CEP para a entrega " & i & ":")
            Set rng = ws.Columns("B").Find(cep, LookIn:=xlValues, LookAt:=xlWhole)
            If Not rng Is Nothing Then
                MsgBox "O CEP já foi cadastrado. Por favor, digite outro CEP.", vbExclamation
            End If
        Loop While Not rng Is Nothing
        
        ' Pede ao usuário a quantidade de clientes neste CEP
        Dim qtdeClientes As Integer
        Do
            qtdeClientes = InputBox("Digite a quantidade de clientes para o CEP " & cep & ":")
            If qtdeClientes <= 0 Then
                MsgBox "A quantidade de clientes deve ser maior que zero.", vbExclamation
            End If
        Loop While qtdeClientes <= 0
        
        ' Preenche os dados na célula correta
        ws.Range("Z3").value = cep
        ws.Range("AG3").value = qtdeClientes
        
        ' Verifica a próxima linha disponível para colar os dados
        Dim linhaAtual As Integer
        linhaAtual = linhaDestino
        Do Until WorksheetFunction.CountA(ws.Range("B" & linhaAtual & ":I" & linhaAtual)) = 0
            linhaAtual = linhaAtual + 1
        Loop
        
        ' Copia os dados para as células corretas sem sobrescrever
        ws.Range("Z3:AG3").Copy
        ws.Range("B" & linhaAtual).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ' Atualiza a linha de destino para a próxima entrega
        linhaDestino = linhaAtual + 1
    Next i
    
    ' Pergunta ao usuário se deseja cadastrar mais entregas
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja cadastrar mais alguma entrega?", vbQuestion + vbYesNo)
    If resposta = vbYes Then
        PreencherEntregas
    Else
        Exit Sub
    End If
    
    Exit Sub

ErrorHandler:
    ' Trata o erro caso o usuário clique em Cancelar
    Exit Sub
End Sub

Sub FormatarColunas()
    ' FormatarColunas Macro
    Columns("B:O").Select
    Columns("B:O").EntireColumn.AutoFit
    Range("B2").Select
End Sub

Sub GuiaCadastro()
    Sheets("Cadastro").Activate
    Range("K1").Select
End Sub

Sub GuiaDadosMelhorias()
    Sheets("Dashboard").Activate
    Range("F47").Select
End Sub

Sub GuiaDashboard()
    Sheets("Dashboard").Activate
    Range("K1").Select
End Sub

Sub GuiaDistancias()
    Sheets("Distâncias").Activate
    Range("K1").Select
End Sub

Sub GuiaSolver()
    Sheets("Solver").Activate
    Range("K1").Select
End Sub

Sub GuiaTempos()
    Sheets("Tempos").Activate
    Range("K1").Select
End Sub

Sub LimparCidades()
    Dim ws As Worksheet
    Dim rngB As Range
    Dim rngC As Range
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir os intervalos das células a serem limpas
    Set rngB = ws.Range("B3:B100000")
    Set rngC = ws.Range("C3:C100000")
    
    ' Limpar o conteúdo das células nos intervalos
    rngB.ClearContents
    rngC.ClearContents
End Sub

Sub LimparCombustivel()
    Dim ws As Worksheet
    Dim rngB As Range
    Dim rngC As Range
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir os intervalos das células a serem limpas
    Set rngB = ws.Range("N3:N1000")
    Set rngC = ws.Range("O3:O1000")
    
    ' Limpar o conteúdo das células nos intervalos
    rngB.ClearContents
    rngC.ClearContents
End Sub

Sub LimparDemandas()
    Dim ws As Worksheet
    Dim rngB As Range
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir os intervalos das células a serem limpas
    Set rngB = ws.Range("D3:D1000")
    
    ' Limpar o conteúdo das células nos intervalos
    rngB.ClearContents
End Sub

Sub LimparEntregas()
    Dim ws As Worksheet
    Dim rngB As Range
    Dim rngC As Range
    Dim rngD As Range
    Dim rngE As Range
    Dim rngF As Range
    Dim rngG As Range
    Dim rngH As Range
    Dim rngI As Range
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir os intervalos das células a serem limpas
    Set rngB = ws.Range("B3:E1000")
    Set rngC = ws.Range("C3:F1000")
    Set rngD = ws.Range("D3:G1000")
    Set rngE = ws.Range("E3:E1000")
    Set rngF = ws.Range("F3:F1000")
    Set rngG = ws.Range("G3:G1000")
    Set rngH = ws.Range("H3:E1000")
    Set rngI = ws.Range("I3:F1000")
    
    ' Limpar o conteúdo das células nos intervalos
    rngB.ClearContents
    rngC.ClearContents
    rngD.ClearContents
    rngE.ClearContents
    rngF.ClearContents
    rngG.ClearContents
    rngH.ClearContents
    rngI.ClearContents
End Sub

Sub LimparProdutos()
    Dim ws As Worksheet
    Dim rngB As Range
    Dim rngC As Range
    Dim rngD As Range
    Dim rngE As Range
    Dim rngF As Range
    Dim rngG As Range
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Definir os intervalos das células a serem limpas
    Set rngB = ws.Range("H3:E1000")
    Set rngC = ws.Range("I3:F1000")
    Set rngD = ws.Range("J3:G1000")
    Set rngE = ws.Range("K3:G1000")
    Set rngF = ws.Range("L3:G1000")
    Set rngG = ws.Range("M3:M1000")
    
    ' Limpar o conteúdo das células nos intervalos
    rngB.ClearContents
    rngC.ClearContents
    rngD.ClearContents
    rngE.ClearContents
    rngF.ClearContents
    rngG.ClearContents
End Sub

Sub FillProducts()
    Dim productQuantity As Variant
    Dim i As Integer
    Dim lastRow As Long
    
    ' Determine the last filled row in column H
    lastRow = Cells(Rows.Count, 8).End(xlUp).row
    
    ' Request the quantity of products from the user
    Do
        productQuantity = Application.InputBox("Enter the quantity of products:", Type:=1)
        
        ' Check if the user canceled the input
        If productQuantity = False Then Exit Sub
        
        ' Check if the quantity of products is valid
        If productQuantity <= 0 Then
            MsgBox "The quantity of products must be greater than zero."
        End If
    Loop While productQuantity <= 0
    
    ' Register the products and fill in the information
    For i = 1 To productQuantity
        Dim productName As Variant
        Dim productWeight As Variant
        Dim boxQuantity As Variant
        
        ' Request the product name
        Do
            productName = Application.InputBox("Enter the name of product " & i & ":", Type:=2)
            
            ' Check if the user canceled the input
            If productName = False Then Exit Sub
            
            ' Check if the product name is valid
            If Len(productName) = 0 Then
                MsgBox "The product name cannot be empty."
            ElseIf CheckExistingProduct(productName, lastRow + i - 2) Then
                MsgBox "The product is already registered. Enter another name."
                productName = ""
            End If
        Loop While Len(productName) = 0
        
        ' Fill in the product name in column H
        Cells(lastRow + i, 8).value = productName
        
        ' Fill in the product code in column I
        Dim productCode As String
        Dim firstLetter As String
        Dim lastLetter As String
        Dim numericCode As String
        
        firstLetter = UCase(Left(productName, 1))
        lastLetter = UCase(Right(productName, 1))
        numericCode = Format(lastRow + i - 2, "0000")
        
        productCode = firstLetter & lastLetter & numericCode
        
        ' Remove the dash ("-") if it exists after the letters
        productCode = Replace(productCode, "-", "")
        
        ' Remove accents from the product code
        productCode = RemoveAccents(productCode)
        
        ' Fill in the product code in column I
        Cells(lastRow + i, 9).value = productCode
        
        ' Request the product weight
        Do
            productWeight = Application.InputBox("Enter the weight of product " & productName & " in grams:", Type:=1)
            
            ' Check if the user canceled the input
            If productWeight = False Then Exit Sub
            
            ' Check if the product weight is valid
            If productWeight <= 0 Then
                MsgBox "The product weight must be greater than zero."
            End If
        Loop While productWeight <= 0
        
        ' Fill in the product weight in column J
        Cells(lastRow + i, 10).value = CDbl(productWeight)
        
        ' Request the quantity per box
        Do
            boxQuantity = Application.InputBox("Enter the quantity of products per box for product " & productName & ":", Type:=1)
            
            ' Check if the user canceled the input
            If boxQuantity = False Then Exit Sub
            
            ' Check if the quantity per box is valid
            If boxQuantity <= 0 Then
                MsgBox "The quantity per box must be greater than zero."
            End If
        Loop While boxQuantity <= 0
        
        ' Fill in the quantity per box in column K
        Cells(lastRow + i, 11).value = CDbl(boxQuantity)
        
        ' Calculate the total weight per box in column L
        Dim totalWeightPerBox As Double
        totalWeightPerBox = CDbl(productWeight) * CDbl(boxQuantity) / 1000
        
        ' Fill in the total weight per box in column L
        Cells(lastRow + i, 12).value = totalWeightPerBox
        
        ' Request the total value of the box
        Dim boxValue As Variant
        Do
            boxValue = Application.InputBox("Enter the total value of the box for product " & productName & ":", Type:=1)
            
            ' Check if the user canceled the input
            If boxValue = False Then Exit Sub
            
            ' Check if the total value of the box is valid
            If boxValue <= 0 Then
                MsgBox "The total value of the box must be greater than zero."
            End If
        Loop While boxValue <= 0
        
        ' Fill in the total value of the box in column M
        Cells(lastRow + i, 13).value = CDbl(boxValue)
    Next i
End Sub

Function CheckExistingProduct(ByVal product As String, ByVal row As Long) As Boolean
    Dim rng As Range
    Dim cel As Range
    
    Set rng = Range("H3:H" & row)
    
    For Each cel In rng
        If cel.value = product Then
            CheckExistingProduct = True
            Exit Function
        End If
    Next cel
    
    CheckExistingProduct = False
End Function

Function RemoveAccents(ByVal text As String) As String
    Dim accents As String
    Dim characters As String
    Dim i As Integer
    
    accents = "ÀÁÂÃÄÅÈÉÊËÌÍÎÏÐÒÓÔÕÖÙÚÛÜÝàáâãäåèéêëìíîïðòóôõöùúûüýÿ"
    characters = "AAAAAAEEEEIIIIDOOOOOUUUUYaaaaaaeeeeiiiiooooouuuuyy"
    
    For i = 1 To Len(accents)
        text = Replace(text, Mid(accents, i, 1), Mid(characters, i, 1))
    Next i
    
    RemoveAccents = text
End Function

Sub SaveWorkbook()
    Application.ScreenUpdating = False
    ThisWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Sub SolverKm()
    ' Reset all Solver settings
    SolverReset
    
    ' Set the objective cell, type of optimization, and variable cells
    SolverOk SetCell:="$Z$8", MaxMinVal:=2, ValueOf:=0, ByChange:="$U$4:$U$12", _
        Engine:=3, EngineDesc:="Evolutionary"
    
    ' Add the constraint that all variable cells must be different
    SolverAdd CellRef:="$U$4:$U$12", Relation:=6, FormulaText:="AllDifferent"
    
    ' Add the constraint that Z6 must equal 1
    SolverAdd CellRef:="$Z$6", Relation:=2, FormulaText:="1"
    
    ' Configure Solver options
    SolverOptions MaxTime:=0, Iterations:=0, Precision:=0.000001, Convergence:=0.01, _
        StepThru:=False, Scaling:=True, AssumeNonNeg:=True, Derivatives:=1
    
    SolverOptions PopulationSize:=50, RandomSeed:=0, Multistart:=False, RequireBounds:=True, _
        MaxSubproblems:=0, MaxIntegerSols:=0, IntTolerance:=0.1, SolveWithout:=False, _
        MaxTimeNoImp:=8
        
    ' Run Solver to find the best route by km
    SolverSolve
    
    ' Select the result cell (Z8)
    Range("Z8").Select
End Sub

Sub SolverHr()
    ' Reset all Solver settings
    SolverReset
    
    ' Set the objective cell, type of optimization, and variable cells
    SolverOk SetCell:="$AH$10", MaxMinVal:=2, ValueOf:=0, ByChange:="$AC$4:$AC$12", _
        Engine:=3, EngineDesc:="Evolutionary"
    
    ' Add the constraint that all variable cells must be different
    SolverAdd CellRef:="$AC$4:$AC$12", Relation:=6, FormulaText:="AllDifferent"
    
    ' Add the constraint that AH6 must equal 1
    SolverAdd CellRef:="$AH$6", Relation:=2, FormulaText:="1"
    
    ' Configure Solver options
    SolverOptions MaxTime:=0, Iterations:=0, Precision:=0.000001, Convergence:=0.01, _
        StepThru:=False, Scaling:=True, AssumeNonNeg:=True, Derivatives:=1
    
    SolverOptions PopulationSize:=50, RandomSeed:=0, Multistart:=False, RequireBounds:=True, _
        MaxSubproblems:=0, MaxIntegerSols:=0, IntTolerance:=0.1, SolveWithout:=False, _
        MaxTimeNoImp:=8
        
    ' Run Solver to find the best route by hours
    SolverSolve
    
    ' Select the result cell (AH10)
    Range("AH10").Select
End Sub

Sub TurnOffScreen()
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

Sub TurnOnScreen()
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
End Sub

Sub FillVehicles()
    Dim vehicleQuantity As Variant
    Dim i As Integer
    Dim lastRow As Long
    
    ' Determine the last filled row in column E
    lastRow = Cells(Rows.Count, 5).End(xlUp).row
    
    ' Request the quantity of vehicles from the user
    Do
        vehicleQuantity = Application.InputBox("Enter the quantity of vehicles:", Type:=1)
        
        ' Check if the user canceled the input
        If vehicleQuantity = False Then Exit Sub
        
        ' Check if the quantity of vehicles is valid
        If vehicleQuantity <= 0 Then
            MsgBox "The quantity of vehicles must be greater than zero."
        End If
    Loop While vehicleQuantity <= 0
    
    ' Fill in the names of the vehicles in column E
    For i = 1 To vehicleQuantity
        Cells(lastRow + i, 5).value = "V" & (lastRow - 2 + i)
    Next i
    
    ' Fill in the types of vehicles in column F
    For i = 1 To vehicleQuantity
        Dim vehicleType As Variant
        
        Do
            vehicleType = Application.InputBox("Enter the type of vehicle " & Cells(lastRow + i, 5).value & ":", Type:=2)
            
            ' Check if the user canceled the input
            If vehicleType = False Then Exit Sub
            
            ' Check if the vehicle type is valid
            If Len(vehicleType) = 0 Then
                MsgBox "The vehicle type cannot be empty."
            End If
        Loop While Len(vehicleType) = 0
        
        Cells(lastRow + i, 6).value = vehicleType
    Next i
    
    ' Fill in the capacities of the vehicles in column G
    For i = 1 To vehicleQuantity
        Dim vehicleCapacity As Variant
        
        Do
            vehicleCapacity = Application.InputBox("Enter the capacity of vehicle " & Cells(lastRow + i, 5).value & " (in kg):", Type:=1)
            
            ' Check if the user canceled the input
            If vehicleCapacity = False Then Exit Sub
            
            ' Check if the vehicle capacity is valid
            If vehicleCapacity <= 0 Then
                MsgBox "The capacity of the vehicle must be greater than zero."
            End If
        Loop While vehicleCapacity <= 0
        
        Cells(lastRow + i, 7).value = CDbl(vehicleCapacity)
    Next i
End Sub

