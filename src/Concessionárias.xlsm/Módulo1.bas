Attribute VB_Name = "Módulo1"
Sub compila()

resposta = MsgBox("Deseja realmente executar a macro?", vbYesNo)

If resposta = 6 Then

    For Each aba In ThisWorkbook.Sheets
    
        If aba.Index > 3 Then
        
        aba.Activate
        
        Range("A2:F100000").ClearContents
        
        
        End If
    
    Next


alon:
    tipo = InputBox("Você deseja compilar os carros novos ou usados?", "Tipo dos Carros", "Novo/Usado")

    If tipo <> "Novo" And tipo <> "Usado" Then
    
    MsgBox ("Favor inserir somente 'Novo' ou 'Usado'")
    
    GoTo alon
    
    End If


    Sheets("Concessionárias").Activate

    For linha = 2 To 9
    
    concessionaria = Cells(linha, 1).Value
    
    Sheets("Resumo").Activate
    
    ActiveSheet.Range("$A$1:$F$1600").AutoFilter Field:=1, Criteria1:=concessionaria
    ActiveSheet.Range("$A$1:$F$1600").AutoFilter Field:=6, Criteria1:=tipo
    
    ult_linha = Range("A1").End(xlDown).Row
    
    Range("A1:F" & ult_linha).Copy
    
    nome_aba = Mid(concessionaria, 7) & " - " & tipo & "s"
    
    Sheets(nome_aba).Activate
    
    Range("A1").PasteSpecial
    
    Sheets("Concessionárias").Activate
    
    Next

Sheets("Resumo").Activate

ActiveSheet.ShowAllData

MsgBox ("Macro executada com sucesso")


End If

End Sub
