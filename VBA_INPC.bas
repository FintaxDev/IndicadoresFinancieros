Attribute VB_Name = "VBA_INEGI"
Option Explicit

Sub Descargar_INPC()
'=============================================================================================
'    Antes de ejecutar la macro, ir a:
'    Herramientas > Referencias... > marcar Microsoft XML, v6.0 > Aceptar
'=============================================================================================
'    Creado por Yair Testas
'=============================================================================================

    Application.ScreenUpdating = False
        
    Dim strTokenINEGI As String
    Dim http As New XMLHTTP60
    Dim objXmlDoc As Object
    Dim avarMonthsArray As Variant
    Dim intColumn As Integer
    Dim intRow As Integer
    Dim objPost As Object
    Dim strDate As String
    Dim intYear As Integer
    Dim intMonth As Integer
    Dim intLastYear As Integer
    
    strTokenINEGI = ""
    
    If strTokenINEGI = "" Then
        GoTo ErrorToken
    End If
    
    With http
        .Open "GET", "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/628194/es/0700/false/BIE/2.0/" & strTokenINEGI & "?type=xml", False
        .send
        Set objXmlDoc = CreateObject("MSXML2.DOMDocument")
        objXmlDoc.LoadXML .responseXML.XML
        If .Status <> 200 Then
            GoTo ErrorConnection
        End If
    End With
    
    ActiveWindow.DisplayGridlines = False
    
    Cells(1, 1).CurrentRegion.Clear
    Cells(1, 1) = "Año/Mes"
    avarMonthsArray = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    For intColumn = 0 To UBound(avarMonthsArray)
        Cells(1, intColumn + 2) = avarMonthsArray(intColumn)
    Next
    
    Cells.Font.Color = RGB(32, 55, 100)
    Cells.VerticalAlignment = xlVAlignCenter
    
    With Cells(1, 1)
        .CurrentRegion.HorizontalAlignment = xlVAlignCenter
        .CurrentRegion.Interior.Color = RGB(32, 55, 100)
        .CurrentRegion.Font.Color = RGB(255, 255, 255)
        .CurrentRegion.Font.Bold = True
    End With
                
    intRow = 1
    For Each objPost In objXmlDoc.SelectNodes("//Series/Serie/OBSERVATIONS/Observation")
        strDate = objPost.SelectNodes(".//TIME_PERIOD")(0).Text
        intYear = Int(Left(strDate, 4))
        intMonth = Int(Right(strDate, 2))
        If intYear <> intLastYear Then
            intRow = intRow + 1
            Cells(intRow, 1) = intYear
        End If
        Cells(intRow, intMonth + 1) = objPost.SelectNodes(".//OBS_VALUE")(0).Text
        intLastYear = intYear
    Next objPost
    
    Range(Cells(2, 1), Cells(intRow, 1)).Font.Bold = True
    Range(Cells(2, 1), Cells(intRow, 1)).HorizontalAlignment = xlVAlignCenter
    
    intRow = 3
    While Cells(intRow, 1) <> ""
        If intRow Mod 2 = 1 Then
            Range(Cells(intRow, 1), Cells(intRow, 13)).Interior.Color = RGB(231, 230, 230)
        End If
        intRow = intRow + 1
    Wend
    
ErrorEnd:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorToken:
    MsgBox "Obtenga su token en https://www.inegi.org.mx/servicios/api_indicadores.html", vbCritical, "Error de token"
    GoTo ErrorEnd

ErrorConnection:
    MsgBox "No se pudo establecer una conexión con la API del INEGI. Error: " & http.statusText, vbCritical, "Error de conexión"
    GoTo ErrorEnd
    
End Sub
