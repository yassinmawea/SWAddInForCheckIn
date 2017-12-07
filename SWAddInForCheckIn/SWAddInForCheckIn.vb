Imports ENOAPILib
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swpublished
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System.Diagnostics.Process
Imports EnoviaSW2Lib
Imports System.Threading
Imports System.Windows.Forms

<Guid("9f23caa0-087f-4f42-b93d-bf02bcd8359c"), ClassInterface(ClassInterfaceType.None), ProgId("SWAddIn - INVENPRO")>
Public Class SWAddInForCheckIn
    Implements IEnoAddIn

    'Dim cts As CancellationTokenSource
    Dim progressBar As Form1
    Dim completeForm As Form2

    Public Sub GetAddInInfo(ByRef poInfo As EnoAddInInfo) Implements IEnoAddIn.GetAddInInfo
        poInfo.mbsAddInName = "SWAddIn - INVENPRO"
        poInfo.mbsCompany = "MAWEA INDUSTRIES"
        poInfo.mbsDescription = "Derive PDF and DXF during checkin"
        poInfo.mlAddInVersion = 1
        poInfo.mlRequiredVersionMajor = EnoLibVer.EnoLibVer_Major
        poInfo.mlRequiredVersionMinor = EnoLibVer.EnoLibVer_Minor
        poInfo.mlRequiredVersionBuild = 0
    End Sub

    Public Sub InsertUIItems(poUI As IEnoUI, eUIComponent As EnoUIComponent, poSelection As EnoSelection) Implements IEnoAddIn.InsertUIItems
        Throw New NotImplementedException()
    End Sub

    Public Sub OnCmd(poCmd As IEnoCmd) Implements IEnoAddIn.OnCmd
        runMacro(poCmd)
        Debug.Print("Sub Ends here")
    End Sub

    ' Invoke JPO to upload PDF and DXF to server
    Sub UploadPDFDXFtoENOVIA(ByVal server As IEnoServer, ByVal item As IEnoSelectionItem)

        Dim file As IEnoFile
        Dim attribs As IEnoAttributeValues
        Dim partNo(1) As String
        Dim rev(1) As String
        Dim jpo As IEnoJPO
        Dim result As String
        Dim parser(2) As String

        Try

            file = server.GetFileFromPath(item.Path)
            Debug.Print("file name is --> " + file.Name)

            attribs = file.GetAttributes()
            partNo(0) = attribs.GetAtt("$$name$$", "@")
            rev(0) = attribs.GetAtt("$$revision$$", "@")

            parser(0) = partNo(0)
            parser(1) = Left(rev(0), InStr(rev(0), ".") - 1)

            Debug.Print("Coming inside UploadPDFDXFtoENOVIA")
            Debug.Print("partNo --> " + partNo(0))
            Debug.Print("rev --> " + parser(1))

            jpo = server.CreateUtility(EnoObjectType.EnoObj_EnoJPO)
            result = jpo.Execute("INV_SWDerivedOutputJPO", "createConnectDerivedOutput", parser)

        Catch e As Exception
            MsgBox(Err.Description)
        End Try

        Exit Sub
    End Sub

    Async Sub runMacro(ByVal poCmd As IEnoCmd)

        Dim t1 As Thread
        Dim t2 As Thread
        'cts = New CancellationTokenSource()
        Try

            t1 = New Thread(AddressOf ProgressMessage)
            t1.SetApartmentState(ApartmentState.STA)
            t1.Start()

            Await Task.Run(Sub() MainProgram(poCmd))
            t1.Abort()
            t1 = Nothing
            Debug.Print("After Abort T1")
            t2 = New Thread(AddressOf CompleteMessage)
            t2.SetApartmentState(ApartmentState.STA)
            Debug.Print("Before start T2")
            t2.Start()
            Debug.Print("After start T2")

        Catch ex As Exception

        End Try

    End Sub

    Sub ProgressMessage(ct As CancellationToken)
        Dim progressBar As Form1

        progressBar = New Form1
        Application.Run(progressBar)
        Thread.Sleep(1000)

        'progressBar.TopMost = True
        'progressBar.DrawingCount(count)
        'progressBar.Open()
    End Sub

    Async Sub CompleteMessage()
        Debug.Print("in Complete Message Thread")

        completeForm = New Form2
        Application.Run(completeForm)
        Thread.Sleep(1000)
        'completeForm.TopMost = True
        'completeForm.Open()

    End Sub

    'Sub WaitFor(NumOfSeconds As Long)
    'Dim SngSec As Long
    'SngSec = Timer + NumOfSeconds

    'Do While Timer < SngSec
    'Threading.Thread.Sleep(500)
    'Loop

    'End Sub

    Sub MainProgram(ByVal poCmd As IEnoCmd)
        On Error Resume Next
        Dim server As IEnoServer
        Dim swApp As Object
        Dim swModel As ModelDoc2
        Dim swModelENO As String
        Dim boolstatusSWStarts As Boolean
        Dim swModelDocExt As ModelDocExtension
        Dim swExportPDFData As ExportPdfData
        Dim swDraw As DrawingDoc
        Dim swPart As PartDoc
        Dim swAssembly As AssemblyDoc
        Dim boolstatus As Boolean
        Dim boolstatusActivateSheet As Boolean
        Dim boolstatusDXF As Boolean
        Dim filenameFull As String
        Dim filenameFullForerver As String
        Dim filenamePDF As String
        Dim filenameDXF As String
        Dim lErrors As Long
        Dim lWarnings As Long
        Dim iErrors As Integer
        Dim iWarnings As Integer
        Dim strSheetPDFName(1) As String
        Dim varSheetPDFName As Array
        Dim strSheetDXFName(1) As String
        Dim varSheetDXFName As Array
        Dim MyPath As String
        Dim MyExt As String
        Dim sel As IEnoSelection
        Dim item As IEnoSelectionItem
        Dim file As IEnoFile
        Dim checkCAD As List(Of String) = New List(Of String)
        Dim partsToClose As New List(Of String)
        Dim partToClose As String
        Dim partToClosePRT As String
        Dim partToCloseASM As String
        Dim bValue As Boolean
        Dim count As Integer
        Dim serverName As String
        Dim view As IEnoServerView
        Dim myProcess As New Process()
        Dim p() As Process
        Dim saveStatus As Boolean
        Dim swNewModel As New ModelDoc2
        Dim result As Integer
        Dim checkinFromExplorer As Boolean
        Dim listPartsAndAssembly As New List(Of IEnoSelectionItem)
        Dim configMgr As ConfigurationManager
        Dim cusPropMgr As CustomPropertyManager
        Dim config As Configuration
        Dim lRetVal As Integer
        Dim enoFile As IEnoFile
        Dim enoFolder As IEnoFolder
        Dim attribs As IEnoAttributeValues
        Dim partNo(1) As String
        Dim vConfigNameArr As Array
        Dim confname As String
        Dim completeMessage As Form2
        Dim syncs As EnoviaSWAddIn
        Dim type As String
        'Dim progressBar As Form1


        ' Retrieving server information
        Debug.Print("Server Name from poCmd --> " + poCmd.Server.Name)
        server = poCmd.Server
        Debug.Print("Server is logged on? --> " & server.IsLoggedIn)

        ' Test to see how many list were in Selection
        sel = poCmd.Selection
        count = sel.Count + 2
        Debug.Print(count)


        ' Check if list only contains document only. If it is, exit sub.
        For Each item In sel
            Dim path As String
            path = item.GetProperty(EnoSelItemProp.Enospi_Path)
            enoFolder = Nothing
            enoFile = server.GetFileFromPath(item.Path, enoFolder)
            type = enoFile.ObjectTypeName
            Debug.Print("ObjectTypeName --> " & enoFile.ObjectTypeName)
            checkCAD.Add(type)
        Next

        If checkCAD.Contains("SW Component Family") Or checkCAD.Contains("SW Assembly Family") Or checkCAD.Contains("SW Drawing") Then

        Else
            Exit Sub
        End If


        'progressBar = New Form1
        'progressBar.TopMost = True
        'progressBar.DrawingCount(count)
        'progressBar.Open()

        ' Open SW if it's not opened yet
        p = Process.GetProcessesByName("SLDWORKS")
        checkinFromExplorer = False
        Debug.Print("P Count --> " & p.Count)
        If p.Count = 0 Then
            checkinFromExplorer = True
            myProcess.StartInfo.FileName = "C:\Program Files\SOLIDWORKS Corp 2017\SOLIDWORKS\SLDWORKS.exe"
            myProcess.Start()
            Threading.Thread.Sleep(10000)
        End If

        'Do
        '  Call WaitFor(5)
        '  Debug.Print(myProcess.MainWindowTitle)
        '  Call WaitFor(1)
        'Loop Until Not myProcess.MainWindowTitle.Equals("splash")

        'swApp = GetObject(, "SldWorks.Application")
        Do
            swApp = Marshal.GetActiveObject("SldWorks.Application.25")
            'Call WaitFor(3)
            Debug.Print("Checking swApp")
        Loop While (swApp Is Nothing)
        swApp.UserControl = False


        Debug.Print("Macro starts")
        'Call WaitFor(3)

        If checkinFromExplorer Then
        Else
            boolstatus = swApp.RunMacro2("C:\Program Files\SolidWorks Corp 2017\swAddInForCheckIn\Sync1.swp", "Sync11", "main", swRunMacroOption_e.swRunMacroUnloadAfterRun, lErrors)
        End If

        Debug.Print("Macro ends")
        'progressBar.IncreaseValue()
        'progressBar.Refresh()

        ' Creating list of parts and assembly for synchronization
        For Each item In sel
            Dim path As String
            path = item.GetProperty(EnoSelItemProp.Enospi_Path)
            enoFolder = Nothing
            enoFile = server.GetFileFromPath(item.Path, enoFolder)
            type = enoFile.ObjectTypeName
            Debug.Print("ObjectTypeName --> " & enoFile.ObjectTypeName)
            Debug.Print("What is choosen ----> " + path)
            filenameFull = Dir(path)
            MyExt = Right(path, 6)
            Dim comp As StringComparison = StringComparison.OrdinalIgnoreCase
            If type.Contains("Component") Then
                If item.GetProperty(EnoSelItemProp.Enospi_CheckIn) = True Then
                    swModel = swApp.OpenDoc6(path, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", iErrors, iWarnings)
                    Debug.Print("what model? -> " & swModel.GetPathName)
                    Debug.Print("Model opened")
                    swApp.QuitDoc("")
                End If
            ElseIf type.Contains("Assembly") Then
                If item.GetProperty(EnoSelItemProp.Enospi_CheckIn) = True Then
                    swAssembly = swApp.OpenDoc6(path, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", iErrors, iWarnings)
                    'swModel = swApp.ActivateDoc3(filenameFull, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, iErrors)
                    swApp.QuitDoc("")
                End If
            End If
        Next
        'progressBar.IncreaseValue()
        'progressBar.Refresh()

        ' Iterating each item in selection to retrieve the path of drawing
        ' and subsequently create the path to temporarily store PDF and DXF
        For Each item In sel
            filenameFullForerver = item.GetProperty(EnoSelItemProp.Enospi_Path)
            enoFolder = Nothing
            enoFile = server.GetFileFromPath(item.Path, enoFolder)
            Debug.Print("ObjectTypeName --> " & enoFile.ObjectTypeName)
            type = enoFile.ObjectTypeName
            Debug.Print("Processing -->" + filenameFullForerver)
            MyExt = Right(filenameFullForerver, 6)                             ' will contain "SLDDRW"
            MyPath = filenameFullForerver
            filenameFull = Dir(filenameFullForerver)                           ' will contain "SolidPart.SLDDRW"
            MyPath = Left(MyPath, InStr(MyPath, filenameFull) - 1)     ' will contain "C:\Folder1\Folder2\"
            filenameFull = Left(filenameFull, InStr(filenameFull, ".") - 1)   ' will contain "SolidPart"
            If type.Contains("Drawing") Then
                If item.GetProperty(EnoSelItemProp.Enospi_CheckIn) = True Then

                    filenamePDF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".pdf"          ' will contain "Default(SolidPart).pdf"
                    filenameDXF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".dxf"          ' will contain "Default(SolidPart).dxf"

                    ' Open drawing, save, close, and reopen for synchronization
                    swDraw = swApp.OpenDoc6(filenameFullForerver, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", iErrors, iWarnings)
                    swModel = swApp.ActivateDoc3(Dir(filenameFullForerver), True, swRebuildOnActivation_e.swRebuildActiveDoc, iErrors)
                    Debug.Print("swDrawing Path -> " + swModel.GetPathName())
                    'swModel = swApp.ActivateDoc3(Dir(filenameFullForerver), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, iErrors)

                    boolstatus = swApp.RunMacro2("C:\Program Files\SolidWorks Corp 2017\swAddInForCheckIn\PDFDXFMacro_Alt.swp", "Personal11", "main", swRunMacroOption_e.swRunMacroUnloadAfterRun, lErrors)

                    ' Invoke JPO to upload PDF and DXF to server
                    UploadPDFDXFtoENOVIA(server, item)

                    ' Increase progress as method is done
                End If
            End If
            'progressBar.IncreaseValue()
            'progressBar.Refresh()

        Next
        'swApp.UserControl = True

        ' If SW was not opened in the first place, then  close SW.
        If checkinFromExplorer = True Then
            myProcess.Kill()
        End If

        'progressBar.CloseForm()
        swApp.UserControl = True

        'stop progressbar
        'progressBar.Close()

    End Sub



End Class