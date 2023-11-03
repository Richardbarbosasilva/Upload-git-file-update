'ultimate script que converte qualquer formato de documento pra qualquer outro formato suportado, abre, edita, fecha e salva qualquer modificação de planilhas excel dentro de uma pasta inteira de documentos!

Sub AutomaçãoFoda()
    
    Dim FolderPathOld As String
    'esta variável identifica a pasta que estão os documentos a serem convertidos para algum outro formato;
    Dim Filename As String
    'esta variável identifica os arquivos dentro de uma pasta que sofrerão a conversão para outra extensão;
    Dim FolderPathNew As String
    'esta variável identifica a pasta onde estão os documentos já convertidos que passarão por algum tratamento de dados da Macro;
    Dim wb As Workbook
    'esta variável é responsável por abrir, salvar e fechar os arquivos;
    Dim FileNameFormatation As String
    'esta variável é responsável por auxiliar a entrada e o fechamento do loop da macro 
    
    
    'Defina o diretório da pasta onde estão os arquivos .csv
    FolderPathOld = CUsersRichard.silvaDownloadsnotacsv
    
    'Especifique o diretório de destino para os arquivos XLSX
    FolderPathNew = CUsersRichard.silvaDownloadsnotaxlsx
    
    'Inicialize a variável Filename
    Filename = Dir(FolderPathOld & .csv)
    
    'Desative atualizações de tela para melhorar o desempenho
    Application.ScreenUpdating = False
    
    'Percorra todos os arquivos .csv na pasta
    Do While Filename  
        'Use Workbooks.OpenText para importar o CSV diretamente como uma nova pasta de trabalho
        Workbooks.OpenText Filename=FolderPathOld & Filename, DataType=xlDelimited, TextQualifier=xlDoubleQuote, ConsecutiveDelimiter=False, Tab=False, Semicolon=True, Comma=False, Space=False, Other=False, Local=True
        
        'Atualize o nome da pasta de trabalho para remover a extensão .csv
        ActiveWorkbook.SaveAs Filename=FolderPathNew & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & .xlsx, FileFormat=51 ' 51 é (pinga) o formato para .xlsx
        ActiveWorkbook.Close SaveChanges=False
        
        'Obtenha o próximo arquivo .csv na pasta
        Filename = Dir
    Loop
    'Ative as atualizações de tela novamente
    
    FileNameFormatation = Dir(FolderPathNew & .xlsx)
    
    
     Do While FileNameFormatation  
        'Abra cada arquivo .xlsx
        Set wb = Workbooks.Open(FolderPathNew & FileNameFormatation)
        
        'Ative a atualização da tela para a execução da macro
        Application.ScreenUpdating = False
        
        'Execute a macro na pasta de trabalho
    Columns(AA).EntireColumn.AutoFit
    Columns(BB).EntireColumn.AutoFit
    Columns(CC).EntireColumn.AutoFit
    Columns(DD).EntireColumn.AutoFit
    Columns(EE).EntireColumn.AutoFit
    Columns(FF).EntireColumn.AutoFit
    Columns(GG).EntireColumn.AutoFit
    Range(A1G1).Select
    Range(G1).Activate
    Selection.Font.Bold = True
    Columns(AG).Select
    Range(G18).Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns(GG).Select
    Selection.Style = Currency
    Selection.ColumnWidth = 15.5
    Columns(CC).Select
    Selection.Style = Currency
      Columns(DD).Select
    Selection.Style = Currency
    Selection.ColumnWidth = 80#
    Columns(CC).Select
    Selection.Style = Currency
    
        
        'Salve as alterações na pasta de trabalho
        wb.Save
        
        'Feche a pasta de trabalho sem perguntar para salvar novamente
        wb.Close SaveChanges=False
        
        'Obtenha o próximo arquivo .xlsx na pasta
        FileNameFormatation = Dir
    Loop
    
'Ative as atualizações de tela novamente
    Application.ScreenUpdating = True
        
End Sub
