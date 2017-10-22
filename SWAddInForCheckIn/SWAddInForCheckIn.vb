Imports ENOAPILib
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swpublished
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System.Diagnostics.Process
Imports swBendLine

<Guid("9f23caa0-087f-4f42-b93d-bf02bcd8359c"), ClassInterface(ClassInterfaceType.None), ProgId("SWAddIn - INVENPRO")>
Public Class SWAddInForCheckIn
    Implements IEnoAddIn


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
        On Error Resume Next

        Dim server As IEnoServer
        Dim swApp As SldWorks
        Dim swModel As ModelDoc2
        Dim boolstatusSWStarts As Boolean
        Dim swModelDocExt As ModelDocExtension
        Dim swExportPDFData As ExportPdfData
        Dim swDraw As DrawingDoc
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
        Dim partsToClose As List(Of String)
        Dim partToClose As String
        Dim partToClosePRT As String
        Dim partToCloseASM As String
        Dim bValue As Boolean
        Dim progressBar As Form1
        Dim count As Integer
        Dim serverName As String
        Dim view As IEnoServerView
        Dim myProcess As New Process()
        Dim p() As Process
        Dim saveStatus As Boolean
        Dim swNewModel As New ModelDoc2
        Dim result As Integer
        Dim checkinFromExplorer As Boolean
        Dim dllobj As New swBendLine.swBendLine
        Dim listPartsAndAssembly As New List(Of IEnoSelectionItem)

        Debug.Print("Server Name from poCmd --> " + poCmd.Server.Name)
        server = poCmd.Server
        Debug.Print("Server is logged on? --> " & server.IsLoggedIn)

        partsToClose = New List(Of String)
        sel = poCmd.Selection
        For Each item In sel
            Dim path As String


            path = item.GetProperty(EnoSelItemProp.Enospi_Path)
            Debug.Print("What is choosen ----> " + path)

            MyExt = Right(path, 6)

            If String.Compare(MyExt, "SLDDRW", True) = 0 Then

                listPartsAndAssembly.Add(item)

            End If

        Next
        count = sel.Count
        Debug.Print(count)

        p = Process.GetProcessesByName("SLDWORKS")
        checkinFromExplorer = False

        Debug.Print("P Count --> " & p.Count)
        If p.Count = 0 Then

            checkinFromExplorer = True
            'myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.FileName = "C:\Program Files\SolidWorks Corp\SOLIDWORKS (2)\SLDWORKS.exe"
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
            myProcess.Start()
            Threading.Thread.Sleep(20000)

        End If
        swApp = GetObject(, "SldWorks.Application")

        If checkinFromExplorer = True Then
            swApp.Visible = False
        End If

        progressBar = New Form1
        progressBar.setMaximum(count)
        progressBar.TopMost = True
        progressBar.Show()

        'Get List of Parts/Assembly and close reopen to sync PartNo and Rev
        For Each item In listPartsAndAssembly
            filenameFullForerver = item.GetProperty(EnoSelItemProp.Enospi_Path)
            swModel = swApp.OpenDoc6(filenameFullForerver, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", iErrors, iWarnings)
            swModel = swApp.ActiveDoc
            saveStatus = swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, iErrors, iWarnings)
            result = swApp.CloseAndReopen(swModel, swCloseReopenOption_e.swCloseReopenOption_DiscardChanges, swNewModel)
            Threading.Thread.Sleep(2000)
        Next

        For Each item In sel
            'Open specified drawing
            'swModel = swApp.ActiveDoc
            'swModelDocExt = swModel.Extension
            'swDraw = swModel
            'swExportPDFData = swApp.GetExportFileData(1)

            'filename = "D:\Documents\Yassin Work\INVENPRO\Testing\PDF Output.PDF"
            'filenameFull = swModel.GetPathName


            filenameFullForerver = item.GetProperty(EnoSelItemProp.Enospi_Path)
            'swApp = GetObject(filenameFullForerver, "SldWorks.Application")
            Debug.Print("Processing -->" + filenameFullForerver)
            MyExt = Right(filenameFullForerver, 6)                             ' will contain "SLDDRW"
            MyPath = filenameFullForerver
            filenameFull = Dir(filenameFullForerver)                           ' will contain "SolidPart.SLDDRW"
            MyPath = Left(MyPath, InStr(MyPath, filenameFull) - 1)     ' will contain "C:\Folder1\Folder2\"
            filenameFull = Left(filenameFull, InStr(filenameFull, ".") - 1)   ' will contain "SolidPart"

            If String.Compare(MyExt, "SLDDRW", True) = 0 Then

                filenamePDF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".pdf"          ' will contain "Default(SolidPart).pdf"
                filenameDXF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".dxf"          ' will contain "Default(SolidPart).dxf"

                boolstatusSWStarts = True
                'boolstatusSWStarts = swApp.StartupProcessCompleted
                'Debug.Print("SW Started? ---->" & boolstatusSWStarts)

                'Do While boolstatusSWStarts = False
                'Threading.Thread.Sleep(3000)
                'swApp = GetObject(, "SldWorks.Application")
                'boolstatusSWStarts = swApp.StartupProcessCompleted
                'Debug.Print("swApp working?? >> " & swApp.Visible)
                'Loop

                'Debug.Print("SW Started Now? ---->" & boolstatusSWStarts)

                If boolstatusSWStarts Then
                    Debug.Print("filenameFullForerver -- > " + filenameFullForerver)
                    swDraw = swApp.OpenDoc6(filenameFullForerver, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", iErrors, iWarnings)

                    Threading.Thread.Sleep(5000)

                    dllobj.CreateBendLine()

                    swModel = swApp.ActiveDoc
                    saveStatus = swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, iErrors, iWarnings)

                    result = swApp.CloseAndReopen(swModel, swCloseReopenOption_e.swCloseReopenOption_DiscardChanges, swNewModel)
                    Threading.Thread.Sleep(5000)

                    swModelDocExt = swNewModel.Extension
                    swExportPDFData = swApp.GetExportFileData(1)
                    ' Names of the sheets
                    strSheetPDFName(0) = "Sheet1"
                    strSheetDXFName(0) = "Sheet2"

                    varSheetPDFName = strSheetPDFName
                    varSheetDXFName = strSheetDXFName


                    If swExportPDFData Is Nothing Then MsgBox("Nothing")
                    boolstatus = swExportPDFData.SetSheets(swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, varSheetPDFName)
                    swExportPDFData.ViewPdfAfterSaving = False

                    boolstatus = swModelDocExt.SaveAs(filenamePDF, 0, 0, swExportPDFData, lErrors, lWarnings)
                    Debug.Print("PDF Saved? --------->" & boolstatus)

                    ' Generate DXF code is here
                    swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfActiveSheetOnly)

                    boolstatusActivateSheet = swDraw.ActivateSheet(strSheetDXFName(0))
                    If boolstatusActivateSheet Then
                        boolstatusDXF = swDraw.SaveAs4(filenameDXF, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, lErrors, lWarnings)
                        Debug.Print("DXF saved? ------sure----> " & boolstatusDXF)
                    End If

                    filenameFull = swNewModel.GetTitle

                    Debug.Print("Title to QuitDoc ---> " + filenameFull)
                    'swApp.QuitDoc(filenameFull)

                    UploadPDFDXFtoENOVIA(server, item)

                    progressBar.increaseProgress()

                End If

            Else
                partsToClose.Add(filenameFull)
                progressBar.increaseProgress()
            End If

        Next

        progressBar.Close()

        If checkinFromExplorer = True Then
            myProcess.Kill()
        End If

    End Sub

    Private Sub UploadPDFDXFtoENOVIA(ByVal server As IEnoServer, ByVal item As IEnoSelectionItem)

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
            result = jpo.Execute("INV_SWCheckJPO", "createConnectDerivedOutput", parser)

        Catch e As Exception
            MsgBox(Err.Description)
        End Try

    End Sub


End Class