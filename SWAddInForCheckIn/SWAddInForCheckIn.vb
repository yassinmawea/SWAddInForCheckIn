Imports ENOAPILib
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swpublished
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System.Diagnostics.Process

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

        partsToClose = New List(Of String)
        sel = poCmd.Selection
        For Each item In sel
            Dim path As String
            path = item.GetProperty(EnoSelItemProp.Enospi_Path)
            Debug.Print("What is choosen ----> " + path)
        Next
        swApp = CreateObject("SldWorks.Application")

        For Each item In sel
            ' Open specified drawing
            'swModel = swApp.ActiveDoc
            'swModelDocExt = swModel.Extension
            'swDraw = swModel
            'swExportPDFData = swApp.GetExportFileData(1)

            ' filename = "D:\Documents\Yassin Work\INVENPRO\Testing\PDF Output.PDF"
            'filenameFull = swModel.GetPathName


            filenameFullForerver = item.GetProperty(EnoSelItemProp.Enospi_Path)
            Debug.Print("Processing -->" + filenameFullForerver)
            MyExt = Right(filenameFullForerver, 6)                             ' will contain "SLDDRW"
            MyPath = filenameFullForerver
            filenameFull = Dir(filenameFullForerver)                           ' will contain "SolidPart.SLDDRW"
            MyPath = Left(MyPath, InStr(MyPath, filenameFull) - 1)     ' will contain "C:\Folder1\Folder2\"
            filenameFull = Left(filenameFull, InStr(filenameFull, ".") - 1)   ' will contain "SolidPart"

            If String.Compare(MyExt, "SLDDRW", True) = 0 Then

                filenamePDF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".pdf"          ' will contain "Default(SolidPart).pdf"
                filenameDXF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".dxf"          ' will contain "Default(SolidPart).dxf"

                boolstatusSWStarts = swApp.StartupProcessCompleted
                Debug.Print("SW Started? ---->" & boolstatusSWStarts)

                Do While boolstatusSWStarts = False
                    Threading.Thread.Sleep(500)
                    boolstatusSWStarts = swApp.StartupProcessCompleted
                Loop

                Debug.Print("SW Started Now? ---->" & boolstatusSWStarts)

                If boolstatusSWStarts Then
                    Debug.Print("filenameFullForerver -- > " + filenameFullForerver)
                    swDraw = swApp.OpenDoc6(filenameFullForerver, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", iErrors, iWarnings)
                    Debug.Print("iErrors ---> " & iErrors)
                    Debug.Print("iWarnings --->" & iWarnings)
                    swModel = swApp.ActiveDoc
                    swModelDocExt = swModel.Extension
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

                    filenameFull = swModel.GetTitle
                    Debug.Print("Title to QuitDoc ---> " + filenameFull)

                    swApp.QuitDoc(filenameFull)

                End If



            Else
                partsToClose.Add(filenameFull)

            End If

        Next

        For Each partToClose In partsToClose
            Debug.Print("Part to close ---- > " + partToClose)
            partToClosePRT = partToClose + ".SLDPRT"
            partToCloseASM = partToClose + ".SLDASM"
            swModel = swApp.ActivateDoc3(partToClosePRT, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, iErrors)
            swApp.QuitDoc("")
            swModel = swApp.ActivateDoc3(partToCloseASM, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, iErrors)
            swApp.QuitDoc("")
        Next
    End Sub
End Class