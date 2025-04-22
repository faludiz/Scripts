' A 25.2-es LibreOffice nem a fájl könyvtárába menti a PDF-et, hanem az utolájára használt könyvtárba
' Ez a szkript a PDF-et a forrás fájl könyvtárába menti

Sub MyExportToPDF
    Dim oDoc As Object
    Dim oFrame As Object
    Dim oDispatch As Object
    Dim oProps(1) As New com.sun.star.beans.PropertyValue
    Dim sFileURL As String
    Dim sDirURL As String
    Dim sDirPath As String
    Dim sFileName As String
    Dim sPDFPath As String
    Dim oShell As Object
    Dim oTranslator As Object
    Dim iPos As Integer

    oDoc = ThisComponent
    If Not oDoc.hasLocation Then
        MsgBox "A dokumentumot előbb el kell menteni!", 48, "Hiba"
        Exit Sub
    End If

    sFileURL = oDoc.URL
    iPos = MyInStrRev(sFileURL, "/")
    sDirURL = Mid(sFileURL, 1, iPos)

    ' Külső fájlrendszer elérési út
    oTranslator = createUnoService("com.sun.star.uri.ExternalUriReferenceTranslator")
    sDirPath = oTranslator.translateToExternal(sDirURL)

    ' PDF fájlnév
    sFileName = GetFileNameWithoutExtension(GetFileNameFromPath(oTranslator.translateToExternal(sFileURL))) & ".pdf"

    ' Export PDF ablak
    oProps(0).Name = "FilterName"
    oProps(0).Value = "writer_pdf_Export"
    oProps(1).Name = "FileName"
    oProps(1).Value = sDirURL & sFileName  ' URL kell!

    oFrame = oDoc.CurrentController.Frame
    oDispatch = createUnoService("com.sun.star.frame.DispatchHelper")
    oDispatch.executeDispatch(oFrame, ".uno:ExportToPDF", "", 0, oProps())

    ' PDF fájl teljes elérési útja
    sPDFPath = sDirPath & sFileName

    ' Fájl megnyitása
    If FileExists(sPDFPath) Then
        oShell = createUnoService("com.sun.star.system.SystemShellExecute")
        oShell.execute("file:///" & Replace(sPDFPath, "\", "/"), "", 0)
    End If
End Sub

' ==== Segédfüggvények ====

Private Function GetFileNameFromPath(sPath As String) As String
    GetFileNameFromPath = Mid(sPath, MyInStrRev(sPath, "/") + 1)
End Function

Private Function GetFileNameWithoutExtension(sFileName As String) As String
    Dim i As Integer
    i = MyInStrRev(sFileName, ".")
    If i > 0 Then
        GetFileNameWithoutExtension = Left(sFileName, i - 1)
    Else
        GetFileNameWithoutExtension = sFileName
    End If
End Function

Private Function FileExists(sPath As String) As Boolean
    Dim oSimpleFileAccess As Object
    oSimpleFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    FileExists = oSimpleFileAccess.exists("file:///" & Replace(sPath, "\", "/"))
End Function

' ==== Saját InStrRev helyettesítő függvény ====
Private Function MyInStrRev(sText As String, sFind As String) As Integer
    Dim i As Integer
    For i = Len(sText) To 1 Step -1
        If Mid(sText, i, Len(sFind)) = sFind Then
            MyInStrRev = i
            Exit Function
        End If
    Next i
    MyInStrRev = 0 ' nem találta meg
End Function
