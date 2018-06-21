Attribute VB_Name = "Módulo1"
Public lPasta As String
Public lNovoArquivo As String

Public Sub lsSepararArquivos()
    frmSeparar.Show
End Sub

Public Sub lsPreparar(ByVal lRng As Range, ByVal lRngBase As Range)
    'Define as classificações
    lRng.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(1, lRngBase.Column), Cells(1048576, lRngBase.Column) _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    'Aplicar classificação
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange lRng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'Realiza a separação das colunas
Public Sub lsSeparar(ByVal lRng As Range, ByVal lRngBase As Range)
    Dim lPastaAtual     As String
    Dim lPlanAtual      As String
    Dim iTotalLinhas    As Long
    Dim iTotalLinhasAux As Long
    Dim i               As Long
    Dim lArquivo        As String
    Dim rngAux          As Range
    Dim lInicio         As Long

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'Guarda a pasta e a planilha ativa
    lPastaAtual = ActiveWorkbook.Name
    lPlanAtual = ActiveSheet.Name
    
    'Identifica a última linha da planilha
    iTotalLinhas = Cells(lRng.Rows.Count, 1).End(xlUp).Row
    
    lInicio = 2
    
    For i = 2 To iTotalLinhas + 1
        'Ativa a planilha da base de dados
        Windows(lPastaAtual).Activate
        ActiveWorkbook.Worksheets(lPlanAtual).Activate
        Set rngAux = Cells(i, lRngBase.Column)
        
        'Cria uma nova planilha
        If rngAux.Value <> lArquivo And lArquivo = "" Then
            ActiveWorkbook.Worksheets(lPlanAtual).Activate
            
            lArquivo = rngAux.Value
            
            lsCriarArquivo lPasta, lArquivo & ".xlsx"
        End If
                
        If rngAux.Value <> lArquivo And lArquivo <> "" Then
            'Copiar
            Windows(lPastaAtual).Activate
            Sheets(lPlanAtual).Select
            Application.CutCopyMode = False
            
            'Copia o cabeçalho
            Rows("1:1").Select
            Selection.Copy
            Windows(lNovoArquivo).Activate
            Range("A1").Select
            ActiveSheet.Paste
            
            'Copia as células
            Windows(lPastaAtual).Activate
            Sheets(lPlanAtual).Select
            
            Range(Cells(lInicio, lRng.Columns(1).Column), Cells(rngAux.Row - 1, lRng.Columns(lRng.Columns.Count).Column)).Select
            'Range("A" & lInicio & ":C" & rngAux.Row - 1).Select
            Selection.Copy
            
            Windows(lNovoArquivo).Activate
            Range("A2").Select
            ActiveSheet.Paste
            
            Cells.Select
            Cells.EntireColumn.AutoFit
            
            lInicio = rngAux.Row
            'Fim copiar
            
            ActiveWorkbook.SaveAs Filename:=lPasta & "\" & lArquivo & ".xlsx", FileFormat:= _
                xlOpenXMLWorkbook, CreateBackup:=False
                
            ActiveWorkbook.Close
            
            'lsGerarCSV lArquivo & " GRS.csv", lPasta
            Application.CutCopyMode = False
            'ActiveWorkbook.Close , vbNo
            
            ActiveWorkbook.Worksheets(lPlanAtual).Activate
            
            lArquivo = rngAux.Value
            
            If lArquivo <> "" Then
                lsCriarArquivo lPasta, lArquivo & ".xlsx"
            End If
        End If
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Windows(lPastaAtual).Activate
    Range("A1").Select
    MsgBox "Planilhas geradas com sucesso!", vbOKOnly, "Guia do Excel"

End Sub

'Criar os arquivos
Private Sub lsCriarArquivo(ByVal lCaminho As String, ByVal lNovo As String)
    Dim lNomeArquivo As String
 
    Workbooks.Add
 
    lNovoArquivo = ActiveWindow.Caption
        
End Sub

'Diálogo para selecionar uma pasta
Public Function SelectFolder() As String

    'Configura o objeto
    Dim fileDialog As fileDialog: Set fileDialog = Excel.Application.fileDialog(msoFileDialogFolderPicker)
    
    'Define as propriedade do objeto
    With fileDialog
        
        'Título da mensagem
        .Title = "Selecione a Pasta"
        
        'Se selecionar alguma pasta, retorna o endereco
        If .Show = True Then       'User pressed action button
            DoEvents
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
    
    Set fileDialog = Nothing
End Function

