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

        Dim swApp As Object
        Dim swModel As ModelDoc2
        Dim boolstatusSWStarts As Boolean
        Dim swModelDocExt As ModelDocExtension
        Dim swExportPDFData As ExportPdfData
        Dim swDraw As DrawingDoc
        Dim boolstatus As Boolean
        Dim boolstatusActivateSheet As Boolean
        Dim boolstatusDXF As Boolean
        Dim filenameFull As String
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

        sel = poCmd.Selection
        For Each item In sel
            Dim path As String
            path = item.GetProperty(EnoSelItemProp.Enospi_Path)
            Debug.Print("What is choosen ----> " + path)
        Next

        For Each item In sel
            ' Open specified drawing
            'swModel = swApp.ActiveDoc
            'swModelDocExt = swModel.Extension
            'swDraw = swModel
            'swExportPDFData = swApp.GetExportFileData(1)

            ' filename = "D:\Documents\Yassin Work\INVENPRO\Testing\PDF Output.PDF"
            'filenameFull = swModel.GetPathName


            filenameFull = item.GetProperty(EnoSelItemProp.Enospi_Path)
            Debug.Print("Processing -->" + filenameFull)
            MyExt = Right(filenameFull, 6)                             ' will contain "SLDDRW"

            If String.Compare(MyExt, "SLDDRW", True) = 0 Then

                swApp = CreateObject("SldWorks.Application")

                MyPath = filenameFull
                filenameFull = Dir(filenameFull)                           ' will contain "SolidPart.SLDDRW"
                MyPath = Left(MyPath, InStr(MyPath, filenameFull) - 1)     ' will contain "C:\Folder1\Folder2\"

                filenameFull = Left(filenameFull, InStr(filenameFull, ".") - 1)   ' will contain "SolidPart"
                filenamePDF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".pdf"          ' will contain "Default(SolidPart).pdf"
                filenameDXF = "\\3dexperience17x\derivedoutput\" & filenameFull & ".dxf"          ' will contain "Default(SolidPart).dxf"

                boolstatusSWStarts = swApp.StartupProcessCompleted
                Debug.Print("SW Started? ---->" + boolstatusSWStarts)

                If boolstatusSWStarts Then

                    swDraw = swApp.OpenDoc6(MyPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", iErrors, iWarnings)
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
                    Debug.Print("PDF Saved? --------->" + boolstatus)

                    ' Generate DXF code is here
                    swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfActiveSheetOnly)

                    boolstatusActivateSheet = swDraw.ActivateSheet(strSheetDXFName(0))
                    If boolstatusActivateSheet Then
                        boolstatusDXF = swDraw.SaveAs4(filenameDXF, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, lErrors, lWarnings)
                        Debug.Print("DXF saved? ------sure----> " & boolstatusDXF)
                    End If

                    swApp.ExitApp()
                swApp = Nothing

                End If
            Else
                Debug.Print("This is not Drawing -->" + filenameFull)

            End If


        Next

    End Sub
End Class