Attribute VB_Name = "Módulo1"
Sub compilacao_concessionarias()

'Confirma se o usuário realmente deseja executar a macro

resposta = MsgBox("Você deseja executar essa macro?", vbYesNo)
    
'Se a resposta for positiva 

If resposta = 6 Then

'Limpa dados antigos

    For Each aba In ThisWorkbook.Sheets
    
        If aba.Index > 3 Then
        
            aba.Activate
            
            Range("A2:F1048576").ClearContents
        
        End If
    
    Next
    
'Verifica se o usuário inseriu os dados corretamente

verificacao_inicial:

'Captura a informação do tipo dos carros

    tipo = InputBox("Você deseja compilar os carros novos ou usados?", "Tipo dos Carros", "Novo/Usado")

    If tipo <> "Novo" And tipo <> "Usado" Then
    
        MsgBox ("Favor inserir somente 'Novo' ou 'Usado'")
    
        GoTo verificacao_inicial
    
    End If
    
'Captura a informação dos nomes das unidades de concessionárias

    Sheets("Concessionárias").Activate

    For linha = 2 To 9
    
        concessionaria = Cells(linha, 1).Value
        
        Sheets("Resumo").Activate
        
'Realiza o filtro
        
        ActiveSheet.Range("$A$1:$F$1600").AutoFilter Field:=1, Criteria1:=concessionaria
        
        ActiveSheet.Range("$A$1:$F$1600").AutoFilter Field:=6, Criteria1:=tipo
        
'Descobre o número da última linha
        
        ult_linha = Range("A1").End(xlDown).Row
        
'Copia as informações
        
        Range("A1:F" & ult_linha).Copy
        
        nome_aba = Mid(concessionaria, 7) & " - " & tipo & "s"

'Seleciona a respectiva aba
        
        Sheets(nome_aba).Activate

'Cola as informações
        
        Range("A1").PasteSpecial
        
        Sheets("Concessionárias").Activate
        
    Next

'Volta para a aba principal

    Sheets("Resumo").Activate
    
'Limpa os filtros
    
    ActiveSheet.ShowAllData
    
'Avisa que a macro foi finalizada
    
    MsgBox ("Macro executada com sucesso")

End If

End Sub
