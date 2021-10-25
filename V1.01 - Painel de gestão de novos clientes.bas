Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call Base_Tratada
    
    If Sheets("BASE TRATADA").Cells(4, 37).Value > 0 Then
        Call Base_Filtrada
        Call Base_Resultados
        Sheets("TD").Select
        Range("B12").Select
        MsgBox ("Ajustar TD de resultados M-1 até D-1")
    Else
        Sheets("HC").Select
        Range("B3").Select
        MsgBox ("Ajustar HC")
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub Base_Tratada()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE INICIAL").Select
    Range("L6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("L6"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("B6").Select

    Sheets("BASE TRATADA").Select
    Range("B5").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C4").Value > 0 Then
        linhaf = linhai - Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C4").Value < 0 Then
        linhaf = linhai + Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B6").Select
    Sheets("BASE INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BASE TRATADA").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    Range("W6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("W7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("W6").Select
    Application.CutCopyMode = False
    Range("B6").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_Filtrada()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE FILTRADA").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE TRATADA").Select
    ActiveSheet.Range("$B$5:$AK$60000").AutoFilter Field:=36, Criteria1:="=1", _
        Operator:=xlAnd
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE FILTRADA").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("BASE TRATADA").Select
    Range("B5").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B6").Select
    Sheets("BASE FILTRADA").Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("AL5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_Resultados()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE RESULTADOS").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE FILTRADA").Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BASE RESULTADOS").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("BASE RESULTADOS").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE RESULTADOS").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B4").Select
    ActiveWorkbook.RefreshAll
    
    Application.ScreenUpdating = True

End Sub

Sub Arquivo_Envio()

    Application.ScreenUpdating = False
    
   ActiveWorkbook.Save
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C13").Value & " - Gestão de Novos Clientes - Dados até dia " & Worksheets("MACROS").Range("C14").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets("PERFORMANCE MoM").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("PERFORMANCE M-1").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("VISÃO GERENCIAL").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(Array("MACROS", "BASE INICIAL", "FECHAMENTO OS", "HC", "BASE TRATADA", _
        "BASE FILTRADA", "TD", "TDP")).Select
    Sheets("BASE FILTRADA").Activate
    Sheets(Array("MACROS", "BASE INICIAL", "FECHAMENTO OS", "HC", "BASE TRATADA", _
        "BASE FILTRADA", "TD", "TDP", "GRÁFICOS")).Select
    Sheets("GRÁFICOS").Activate
    Application.CutCopyMode = False
    ActiveWorkbook.Connections( _
        "WorksheetConnection_BASE RESULTADOS MoM!$B$3:$AC$2567").Delete
    ActiveWindow.SelectedSheets.Delete
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("VISÃO GERENCIAL").Select
    Range("B6").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("PERFORMANCE M-1").Select
    Range("B5").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("PERFORMANCE MoM").Select
    Range("B6").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True

End Sub

