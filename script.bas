Private Sub Workbook_Open()
 
' Mensagem inicial que solicta o nome > EstaPastaDeTrabalho

    If Sheets(39).Cells(3, 2).Value = "" Then
        Sheets("Idioma").Activate
        Idioma.Show
    ElseIf Sheets(39).Cells(3, 2).Value = "Português" Then
        Sheets("Geral").Activate
    ElseIf Sheets(39).Cells(3, 2).Value = "English" Then
        Sheets("Overview").Activate
    End If

    If Sheets(39).Cells(3, 2).Value = "Português" Then
        Sheets("Geral").Activate
        If Cells(4, 11).Value = "{nome}" Then
            Boas_Vindas.Show
        End If
    ElseIf Sheets(39).Cells(3, 2).Value = "English" Then
        Sheets("Overview").Activate
        If Cells(4, 11).Value = "{name}" Then
            Welcome.Show
        End If
    End If
    
ThisWorkbook.Save
            
End Sub

Sub Alterar_idioma()

' Alterar idioma

    Sheets("Idioma").Activate
    Idioma.Show
    If Sheets(39).Cells(3, 2).Value = "Português" Then
        Sheets("Geral").Activate
    ElseIf Sheets(1).Cells(3, 2).Value = "English" Then
        Sheets("Overview").Activate
    End If

    If Sheets(39).Cells(3, 2).Value = "Português" Then
        Sheets("Geral").Activate
        If Cells(4, 11).Value = "{nome}" Then
            Boas_Vindas.Show
        End If
    ElseIf Sheets(39).Cells(3, 2).Value = "English" Then
        Sheets("Overview").Activate
        If Cells(4, 11).Value = "{name}" Then
            Welcome.Show
        End If
    End If
    
End Sub

Sub Atualizar_Grafico()

' Atualizar gráfico de cada mês

ThisWorkbook.RefreshAll

If Cells(17, 2).Value = "ENTRADAS" Then

With ActiveSheet.PivotTables("Tabela dinâmica13").PivotFields("CATEGORIA")
        .ClearAllFilters
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
End With
        
ElseIf Cells(17, 2).Value = "INCOMES" Then

With ActiveSheet.PivotTables("Tabela dinâmica13").PivotFields("CATEGORY")
        .ClearAllFilters
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
End With

End If

End Sub

Sub Novo_Registro_Botao()

'Botão de Novo Registro

If Cells(17, 2).Value = "ENTRADAS" Then
Novo_Registro.Show
ElseIf Cells(17, 2).Value = "INCOMES" Then
Novo_Registro_En.Show
End If

End Sub

Sub Inserir()
'
' Botão "Inserir"
'
'
    If Cells(11, 5).Value = "Fonte" Or Cells(11, 5).Value = "Source" Then
    Range("E12:F12").Select
    Selection.Copy
    Range("B17").Select
    
    Inserir_Fonte
    
    ElseIf Cells(11, 5).Value = "Item" Then
    Range("E12:G12").Select
    Selection.Copy
    Range("E17").Select

    Inserir_Item

    End If
    
    Application.CutCopyMode = False
    Range("E12").Select
    
End Sub



Sub Inserir_Fonte()

' Inserir uma entrada

    If Cells(18, 2).Value <> "" Then
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Else
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    End If
    
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
End Sub

Sub Inserir_Item()

' Inserir uma saída

    If Cells(18, 5).Value <> "" Then
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Else
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    End If
    
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G12").Select
    ActiveCell.FormulaR1C1 = ""
End Sub

Sub Excluir_Ultimo_Botao()

' Botão para excluir último registro

If Cells(17, 2).Value = "ENTRADAS" Then
Excluir_Ultimo.Show
ElseIf Cells(17, 2).Value = "INCOMES" Then
Excluir_Ultimo_En.Show
End If

End Sub

Sub Excluir_Entrada()

' Excluir o último registro se o registro é uma entrada

    Range("B19").Select
    Selection.End(xlDown).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""

End Sub

Sub Excluir_Saida()

' Excluir o último registro se o registro é uma saída

    Range("E19").Select
    Selection.End(xlDown).Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""

End Sub

Sub Limpar_Lista()

' Limpar Lista de Compras

    Range("H16:J16").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveCell.Range("A1:B1").Select
End Sub


Sub Gastos_por_categorias()

' Busca os produtos de acordo com a categoria selecionada

Application.ScreenUpdating = False
Application.DisplayAlerts = False

linha = 20
For i = 8 To 19
While Worksheets(i).Cells(linha, 6) <> ""

        If Worksheets(i).Cells(linha, 6).Value = Worksheets(6).Cells(15, 10).Value Then
        Worksheets(i).Select
        Worksheets(i).Cells(linha, 5).Select
        ActiveCell.Range("A1:C1").Select
        Selection.Copy
        Condicao
            If Cells(16, 7).Value = "" Then
                Cells(16, 7).Select
                ActiveCell.Value = Worksheets(i).Name
            Else
                Cells(15, 7).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Range("A1").Select
                ActiveCell.Value = Worksheets(i).Name
            End If
        linha = linha + 1
        
        Else
        linha = linha + 1
        
        End If
Wend
linha = 20
Next
    
End Sub

Sub Condicao()

' Condição para a macro Gastos_por_categorias()

    Sheets("Gastos por Categorias").Select
    
    If Cells(16, 4).Value = "" Then
        Cells(16, 4).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Else
        Cells(15, 4).Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
End Sub
 
Sub Limpar_Gastos()

' Limpar os produtos na aba de categorias

    Application.ScreenUpdating = False

If Cells(16, 5).Value <> "" Then
    Cells(16, 4).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Cells(16, 4).Select
End If
    
End Sub

Sub New_Entrada()
' Quando selecionada a opção de "Entrada" em "Criar Novo Registro"

' Limpar campos 

    Range("E11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
	
' Formatar

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    Range("G11").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "Fonte"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "Preço"
    Range("E11:F12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    With Selection.Font
        .Name = "Segoe UI Semibold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("G10:G13").Select
    Range("G11:G12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("E12:G12").Select
    With Selection.Font
        .Name = "Segoe UI"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("F12").Select
    Selection.NumberFormat = "$#,##0.00"
    Range("E12").Select
    End With
End Sub

Sub New_Saida()

'  Quando selecionada a opção de "Saída" em "Criar Novo Registro"

' Limpar campos

    Range("E11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G12").Select
    ActiveCell.FormulaR1C1 = ""
	
' Formatar 

    Range("G11").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "Item"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "Categoria"
    Range("G11").Select
    ActiveCell.FormulaR1C1 = "Preço"
    Range("E11:G11").Select
    With Selection.Font
        .Name = "Segoe UI Semibold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("E12:G12").Select
    With Selection.Font
        .Name = "Segoe UI"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("F12").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Categorias!$H$15:$H$41"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("G12").Select
    Selection.NumberFormat = "$#,##0.00"
        Range("E11:F12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E12").Select
End Sub

' Tutoriais abaixo

Sub Tutorial_Categorias()

MsgBox ("Aqui você pode buscar quais produtos comprou de acordo com cada categoria. Para isso, selecione uma categoria em ""Setor"" e clique em ""Buscar"".")
End Sub

Sub Tutorial_Metas()

MsgBox ("Nessa aba, você pode estabelecer metas para economizar em cada mês. O valor poupado é calculado automaticamente")
End Sub

Sub Tutorial_Editar()

MsgBox ("Aqui você pode editar as categorias de acordo com o que você consume. Sinta-se livre para alterar como preferir.")

End Sub
Sub Tutorial_Lista()

MsgBox ("Esta é sua lista de compras! Anote aqui suas compras para a próxima ida ao mercado, e, quando feito, basta clicar em ""Limpar Lista"" para poder criar uma nova.")
End Sub

Sub Tutorial_Mes()

MsgBox ("Para registrar uma nova entrada ou saída, clique em ""Criar novo Registro"". Ao aparecer o campo, preencha as informações necessárias e clique em ""Inserir"".")
End Sub

' Versões em Inglês abaixo

Sub New_Saida_En()

'  Quando selecionada a opção de "Saída" em "Criar Novo Registro"

    Range("E11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G11").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "Item"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("G11").Select
    ActiveCell.FormulaR1C1 = "Price"
    Range("E11:G11").Select
    With Selection.Font
        .Name = "Segoe UI Semibold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("E12:G12").Select
    With Selection.Font
        .Name = "Segoe UI"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("F12").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Categories!$H$15:$H$41"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("G12").Select
    Selection.NumberFormat = "$#,##0.00"
        Range("E11:F12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E12").Select
End Sub


Sub New_Entrada_En()

' Quando selecionada a opção de "Entrada" em "Criar Novo Registro"

    Range("E11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F11").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G12").Select
    ActiveCell.FormulaR1C1 = ""
    Range("F12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    Range("G11").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "Source"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "Price"
    Range("E11:F12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    With Selection.Font
        .Name = "Segoe UI Semibold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("G10:G13").Select
    Range("G11:G12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("E12:G12").Select
    With Selection.Font
        .Name = "Segoe UI"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .ThemeFont = xlThemeFontNone
    End With
    Range("F12").Select
    Selection.NumberFormat = "$#,##0.00"
    Range("E12").Select
    End With
End Sub


Sub Gastos_por_categorias_En()

' Busca os produtos de acordo com a categoria selecionada

Application.ScreenUpdating = False
Application.DisplayAlerts = False

linha = 20
For i = 27 To 38
While Worksheets(i).Cells(linha, 6) <> ""

        If Worksheets(i).Cells(linha, 6).Value = Worksheets(25).Cells(15, 10).Value Then
        Worksheets(i).Select
        Worksheets(i).Cells(linha, 5).Select
        ActiveCell.Range("A1:C1").Select
        Selection.Copy
        Condicao_En
            If Cells(16, 7).Value = "" Then
                Cells(16, 7).Select
                ActiveCell.Value = Worksheets(i).Name
            Else
                Cells(15, 7).Select
                Selection.End(xlDown).Select
                ActiveCell.Offset(1, 0).Range("A1").Select
                ActiveCell.Value = Worksheets(i).Name
            End If
        linha = linha + 1
        
        Else
        linha = linha + 1
        
        End If
Wend
linha = 20
Next
    
End Sub

Sub Condicao_En()

' Condição para a macro Gastos_por_categorias()

    Sheets("Expenses by Category").Select
    
    If Cells(16, 4).Value = "" Then
        Cells(16, 4).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Else
        Cells(15, 4).Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
End Sub

Sub Tutorial_Lista_En()

MsgBox ("This is your shopping list! Make a note of your purchases here for your next trip to the market, and when done, just click on ""Reset List"" to create a new one.")
End Sub

Sub Tutorial_Editar_En()

MsgBox ("Here you can edit the categories according to what you buy most. Feel free to change as you like.")

End Sub

Sub Tutorial_Metas_En()

MsgBox ("In this tab, you can set goals to save each month. The amount saved is calculated automatically.")
End Sub

Sub Tutorial_Categorias_En()

MsgBox ("Here you can search for which products you purchased according to each category. To do this, select a category under ""Sector"" and click on ""Search"".")
End Sub

Sub Tutorial_Mes_En()

MsgBox ("To register a new Income or Expense, click on ""Create New Entry"". When the field appears, complete the necessary information and then click on ""Insert"".")
End Sub

Sub Gravar()

' Salva a planilha automáticamente a cada minuto

     ThisWorkbook.Save
     Call Timer
End Sub

Sub Timer()

' Salva a planilha automáticamente a cada minuto

     Application.OnTime Now + TimeValue("00:01:00"), "gravar"
    
End Sub

' Formulários abaixo

Private Sub CommandButton2_Click()

' Botão de sair

ActiveWorkbook.Close SaveChanges:=True

End Sub


Private Sub Ok_botao_Click()

' Botão de "Ok" na mensagem inicial
' Altera o nome para o nome inserido

    Boas_Vindas.Hide
    
    Application.ScreenUpdating = False
    
    For i = 1 To 7
    ActiveWorkbook.Sheets(i).Activate
    Cells(4, 11).Value = Boas_Vindas.TextBox1.Value
    Next
    
    MsgBox ("Muito obrigada, " + Boas_Vindas.TextBox1.Value + "!")
    ActiveWorkbook.Sheets(1).Activate

End Sub

Private Sub Entrada2_Click()

' Excluir último - verificar se é entrada 

    Unload Excluir_Ultimo
    Excluir_Entrada
End Sub

Private Sub Saida2_Click()

' Excluir último - verificar se é saída 

    Unload Excluir_Ultimo
    Excluir_Saida

End Sub


Private Sub CommandButton1_Click()

' Linguagem escolhida = "Inglês"

Idioma.Hide
Cells(3, 2).Value = "English"

End Sub

Private Sub CommandButton3_Click()

' Linguagem escolhida = "Português"

Idioma.Hide
Cells(3, 2).Value = "Português"

End Sub



Private Sub Entrada1_Click()

' Novo registro - verificar se é entrada 

Unload Novo_Registro
New_Entrada

End Sub

Private Sub Saida1_Click()

' Novo registro - verificar se é saída 

Unload Novo_Registro
New_Saida

End Sub


Private Sub Entrada_Click()

' Novo registro - verificar se é entrada - versão em inglês

Unload Novo_Registro_En
New_Entrada_En

End Sub

Private Sub Saida_Click()

' Novo registro - verificar se é saída - versão em inglês

Unload Novo_Registro_En
New_Saida_En

End Sub


Private Sub CommandButton2_Click()

' Botão de sair

ActiveWorkbook.Close SaveChanges:=True

End Sub

Private Sub Ok_botao_Click()

' Botão de "Ok" na mensagem inicial
' Altera o nome para o nome inserido

    Welcome.Hide
    
    Application.ScreenUpdating = False
    
    For i = 20 To 26
    ActiveWorkbook.Sheets(i).Activate
    Cells(4, 11).Value = Welcome.TextBox1.Value
    Next
    
    MsgBox ("Thank you, " + Welcome.TextBox1.Value + "!")
    ActiveWorkbook.Sheets(20).Activate
    
End Sub





