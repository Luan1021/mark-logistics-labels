Attribute VB_Name = "Módulo1"

Sub ImprimirEtiquetas()
    ProcessarEtiquetas False
End Sub

Sub VisualizarEtiquetas()
    ProcessarEtiquetas True
End Sub

Private Sub ProcessarEtiquetas(ByVal soVisualizar As Boolean)

    Dim wsDados As Worksheet
    Dim wsEt As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long, v As Long
    Dim cargaEscolhida As String
    Dim cargaAtual As String
    Dim contadorVolTotal As Long
    Dim qtdVol As Long
    
    Set wsDados = ThisWorkbook.Worksheets("DADOS")
    Set wsEt = ThisWorkbook.Worksheets("ETIQUETA")
    
    
    cargaEscolhida = InputBox("Digite o número da CARGA que deseja usar:", _
                              IIf(soVisualizar, "Visualizar etiquetas", "Imprimir etiquetas"))
    If cargaEscolhida = "" Then
        If MsgBox("Nenhuma carga digitada. Deseja realmente usar TODAS as cargas da aba DADOS?", _
                  vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    
   
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    
    cargaAtual = ""
    contadorVolTotal = 0
    
    For i = 2 To ultimaLinha
        
        Dim cargaLinha As String
        Dim cod As String
        Dim descProd As String
        Dim transp As String
        
        cargaLinha = wsDados.Cells(i, 1).Text
        cod = wsDados.Cells(i, 2).Text
        descProd = wsDados.Cells(i, 3).Text
        transp = wsDados.Cells(i, 5).Text
        qtdVol = CLng(wsDados.Cells(i, 4).Value)
        
        If qtdVol < 1 Then qtdVol = 1
        
        
        If cargaEscolhida = "" Or cargaLinha = cargaEscolhida Then
            
            
            If cargaLinha <> cargaAtual Then
                cargaAtual = cargaLinha
                contadorVolTotal = 0
            End If
            
            
            For v = 1 To qtdVol
                
                contadorVolTotal = contadorVolTotal + 1
                
               
                wsEt.Range("B1").Value = "'" & cargaLinha
                wsEt.Range("B2").Value = "'" & cod
                wsEt.Range("A3").Value = descProd
                wsEt.Range("B4").NumberFormat = "@"
                wsEt.Range("B4").Value = "'" & CStr(v) & "/" & CStr(qtdVol)
                wsEt.Range("B5").Value = contadorVolTotal
                wsEt.Range("B6").Value = transp
                
                DoEvents
                
                If soVisualizar Then
                    wsEt.PrintPreview
                Else
                Dim nomeImpressora As String
                Application.Dialogs(xlDialogPrinterSetup).Show
                nomeImpressora = Application.ActivePrinter
                Application.ActivePrinter = nomeImpressora
                wsEt.PrintOut Copies:=1
                End If
                
            Next v
        End If
    Next i
    
    If soVisualizar Then
        MsgBox "Visualização concluída.", vbInformation
    Else
        MsgBox "Impressão concluída.", vbInformation
    End If
    
End Sub


