Option Strict Off
Option Explicit On
Imports System.Math
Imports System.Data.OleDb
Imports System.Data
Imports System.Data.Common
Imports System.Windows.Forms
Imports SAPbouiCOM.BoAppEventTypes
Imports Prj_AN.clsGlobal


Imports Microsoft.Office.Interop.Excel


Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System
Imports System.Threading
Imports System.Security.Permissions
Imports System.Management


'Imports Excel


Public Class clsMain    

    Public oCompany As SAPbobsCOM.Company
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter
    'Private oForm As SAPbouiCOM.Form


    Dim objSBOGuiApi As SAPbouiCOM.SboGuiApi
    Dim ClsGlobal As New clsGlobal

    Const TBar_Update As String = "Update Status"    'toolbar Update Status Project Master
    Const TBar_ClearRow As String = "Clear Row"
    'Const TBar_Delete As String = "Delete"
    Const Tbar_Data As String = "1280"
    Const TBar_LockScr As String = "524"    'toolbar Lock screen
    Const TBar_Find As String = "1281"  'toolbar find
    Const TBar_Add As String = "1282"   'toolbar add
    Const TBar_First As String = "1290"   'toolbar first
    Const TBar_Next As String = "1288"   'toolbar next
    Const TBar_Prev As String = "1289"   'toolbar prev
    Const TBar_Last As String = "1291"   'toolbar last
    Const TBar_Remove As String = "Delete"    'toolbar remove
    Const TBar_FilterTbl As String = "4870" 'toolbar Filter table
    Const TBar_Close As String = "1286"   'toolbar add
    Const TBar_Duplicate As String = "1287"   'toolbar add
    Const TBar_CloseRow As String = "1299"   'toolbar close row --> used by sbo2005a sp:01
    Const TBar_PLDesigner As String = "5895"    'toolbar print layout designer
    Const TBar_FSetting As String = "5890"  'toolbar form setting
    Const TBar_QManager As String = "4865"  'toolbar queries manager
    Const TBar_PrintPrv As String = "519"    'toolbar Print Preview

    ' Item Master Data
    Const ItemMaster_MenuId As String = "3073"
    Const ItemMaster_FormId As String = "150"
    Dim objFormItemMaster As SAPbouiCOM.Form
    Dim intRowItemMaster As Integer

    Const ItemMasterUDF_FormId As String = "-150"
    Dim objFormItemMasterUDF As SAPbouiCOM.Form
    Dim intRowItemMasterUDF As Integer

    ' GOOD RECEIPT
    Const GoodReceipt_MenuId As String = "3078"
    Const GoodReceipt_FormId As String = "721"
    Dim objFormGoodReceipt As SAPbouiCOM.Form
    Dim intRowGoodReceiptDetail As Integer

    Const GoodReceiptUDF_FormId As String = "-721"
    Dim objFormGoodReceiptUDF As SAPbouiCOM.Form
    Dim intRowGoodReceiptDetailUDF As Integer

    Const ListGI_FormId As String = "2000000002"
    Dim intFormCountListGI As Integer
    Dim strCurntListGI As String
    Dim blnModalListGI As Boolean
    Dim objFormListGI As SAPbouiCOM.Form

    ' GOOD ISSUE
    Const GoodIssue_MenuId As String = "3079"
    Const GoodIssue_FormId As String = "720"
    Dim objFormGoodIssue As SAPbouiCOM.Form
    Dim intRowGoodIssueDetail As Integer

    Const GoodIssueUDF_FormId As String = "-720"
    Dim objFormGoodIssueUDF As SAPbouiCOM.Form
    Dim intRowGoodIssueDetailUDF As Integer

    Const ListGR_FormId As String = "2000000001"
    Dim intFormCountListGR As Integer
    Dim strCurntListGR As String
    Dim blnModalListGr As Boolean
    Dim objFormListGR As SAPbouiCOM.Form

    ' INVENTORY TRANSFER
    Const InvTransfer_MenuId As String = "3080"
    Const InvTransfer_FormId As String = "940"
    Dim objFormInvTransfer As SAPbouiCOM.Form
    Dim intRowInvTransferDetail As Integer

    Const InvTransferUDF_FormId As String = "-940"
    Dim objFormInvTransferUDF As SAPbouiCOM.Form
    Dim intRowInvTransferDetailUDF As Integer

    Const ListInvTransfer_FormId As String = "2000000010"
    Dim intFormCountListInvTransfer As Integer
    Dim strCurntListInvTransfer As String
    Dim blnModalListInvTransfer As Boolean
    Dim objFormListInvTransfer As SAPbouiCOM.Form

    ' Batch Master
    Const Batch_FormId As String = "41"
    Dim objFormBatch As SAPbouiCOM.Form
    Dim intRowBatch As Integer

    'Project Master
    Const pictCFL As String = "CFL.bmp"
    Public Const Inventory_MenuId As String = "3072"

    Public Const ProjectMaster_MenuId As String = "30721"
    Public Const ProjectMaster_MenuDesc As String = "Project Master"
    Public Const ProjectMaster_FormId As String = "2000030721"
    Dim intFormCountProjectMaster As Integer
    Dim objFormProjectMaster As SAPbouiCOM.Form

    'Project HARVEST

    Public Const ProjectHarvest_MenuId As String = "30722"
    Public Const ProjectHarvest_MenuDesc As String = "Project Harvesting"
    Public Const ProjectHarvest_FormId As String = "2000030722"
    Dim intFormCountProjectHarvest As Integer
    Dim oFormProjectHarvest As SAPbouiCOM.Form



    ' lookup Net
    Const LookUp_FormId As String = "2000000003"    'look up for general purpose
    Dim intFormCountLookUp As Integer
    Dim blnModalLookUp As Boolean
    Dim strCurntLookUp As String
    Dim intCurntRowLookUp As Integer
    Dim oFormLookUp As SAPbouiCOM.Form

    ' lookup Net Harvest
    Const LookUpNet_FormId As String = "2000000006"    'look up for general purpose
    Dim intFormCountLookUpNet As Integer
    Dim blnModalLookUpNet As Boolean
    Dim strCurntLookUpNet As String
    Dim intCurntRowLookUpNet As Integer
    Dim oFormLookUpNet As SAPbouiCOM.Form

    ' lookup Species
    Const LookUpSpecies_FormId As String = "2000000004"    'look up for general purpose
    Dim intFormCountLookUpSpecies As Integer
    Dim blnModalLookUpSpecies As Boolean
    Dim strCurntLookUpSpecies As String
    Dim intCurntRowLookUpSpecies As Integer
    Dim oFormLookUpSpecies As SAPbouiCOM.Form

    ' lookup Batch
    Const LookUpBatch_FormId As String = "2000000005"    'look up for general purpose
    Dim intFormCountLookUpBatch As Integer
    Dim blnModalLookUpBatch As Boolean
    Dim strCurntLookUpBatch As String
    Dim intCurntRowLookUpBatch As Integer
    Dim oFormLookUpBatch As SAPbouiCOM.Form

    ' lookup Batch
    Const LookUpDistNumber_FormId As String = "2000000007"    'look up for general purpose
    Dim intFormCountLookUpDistNumber As Integer
    Dim blnModalLookUpDistNumber As Boolean
    Dim strCurntLookUpDistNumber As String
    Dim intCurntRowLookUpDistNumber As Integer
    Dim oFormLookUpDistNumber As SAPbouiCOM.Form

    ' Batch RM
    Const BatchRM_FormId As String = "2000000008"    'look up for Batch Raw Material
    Dim intFormCountBatchRM As Integer
    Dim blnModalLookBatchRM As Boolean
    Dim strCurntLookBatchRM As String
    Dim intCurntRowLookBatchRM As Integer
    Dim oFormBatchRM As SAPbouiCOM.Form

    ' Batch FG
    Const BatchFG_FormId As String = "2000000009"    'look up for Batch Raw Material
    Dim intFormCountBatchFG As Integer
    Dim blnModalLookBatchFG As Boolean
    Dim strCurntLookBatchFG As String
    Dim intCurntRowLookBatchFG As Integer
    Dim oFormBatchFG As SAPbouiCOM.Form

    ' lookup Mortalitiy
    Const LookUpMortal_FormId As String = "2000000011"
    Dim intFormCountLookUpMortal As Integer
    Dim blnModalLookUpMortal As Boolean
    Dim strCurntLookUpMortal As String
    Dim intCurntRowLookUpMortal As Integer
    Dim oFormLookUpMortal As SAPbouiCOM.Form

    Dim FileName As String
    Dim Row As Integer
    Dim Col As Integer


    Public Function FindFile() As String

        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()

            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While
            If FileName <> "" Then
                Return FileName
            End If
        Catch ex As Exception
            objApplication.MessageBox("FindFile" & ex.Message)
        End Try

        Return ""

    End Function


    Private Function GetProcessUserName(ByVal Process As Process) As String
        Dim sq As New ObjectQuery("Select * from Win32_Process Where ProcessID = '" & Process.Id & "'")
        Dim searcher As New ManagementObjectSearcher(sq)


        If searcher.Get.Count = 0 Then Return Nothing

        For Each oReturn As ManagementObject In searcher.Get
            Dim o As String() = New String(1) {}

            'Invoke the method and populate the o var with the user name and domain                         
            oReturn.InvokeMethod("GetOwner", DirectCast(o, Object()))

            Return o(0)
        Next
        'Return ""
    End Function


    Public Sub ShowFolderBrowser()

        Dim MyProcs() As System.Diagnostics.Process
        Dim UserName = Environment.UserName

        FileName = ""
        Dim OpenFile As New OpenFileDialog

        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "All files(*.)|*.*"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try

            OpenFile.FilterIndex = filterindex

            OpenFile.RestoreDirectory = True
            MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One")


            For i As Integer = 0 To UBound(MyProcs)
                If GetProcessUserName(MyProcs(i)) = UserName Then
                    GoTo NEXT_STEP
                End If
            Next
            objApplication.MessageBox("Unable to determine Running processes by UserName!")
            OpenFile.Dispose()
            GC.Collect()
            Exit Sub
NEXT_STEP:
            If MyProcs.Length = 1 Then
                'For i As Integer = 0 To MyProcs.Length - 1
                Dim MyWindow As New WindowWrapper(MyProcs(0).MainWindowHandle)
                Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)

                If ret = DialogResult.OK Then
                    FileName = OpenFile.FileName
                    OpenFile.Dispose()
                Else
                    System.Windows.Forms.Application.ExitThread()
                End If
                'Next
            ElseIf MyProcs.Length = 2 Then
                Dim MyWindow As New WindowWrapper(MyProcs(1).MainWindowHandle)
                Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)

                If ret = DialogResult.OK Then
                    FileName = OpenFile.FileName
                    OpenFile.Dispose()
                Else
                    System.Windows.Forms.Application.ExitThread()
                End If
                'objApplication.MessageBox("More than 1 SAP B1 is started!")
            End If
            'objFormGoodIssue.Items.Item("txtGen").Specific.string = FileName
            objFormGoodIssue.Items.Item("btnGen").Specific.Caption = "Upload"
        Catch ex As Exception
            objApplication.StatusBar.SetText(ex.Message)
            FileName = ""
        Finally

            OpenFile.Dispose()
            GC.Collect()
        End Try

    End Sub

    'Public Sub ShowFolderBrowser()

    '    Dim MyProcs() As System.Diagnostics.Process
    '    Dim UserName = Environment.UserName
    '    FileName = ""
    '    Dim OpenFile As New OpenFileDialog

    '    Try
    '        OpenFile.Multiselect = False
    '        OpenFile.InitialDirectory = "C:\"
    '        OpenFile.Filter = "All files(*.)|*.*"
    '        Dim filterindex As Integer = 0
    '        Try
    '            filterindex = 0
    '        Catch ex As Exception
    '        End Try

    '        OpenFile.FilterIndex = filterindex

    '        OpenFile.RestoreDirectory = True
    '        MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One")

    '        If MyProcs.Length = 1 Then
    '            For i As Integer = 0 To MyProcs.Length - 1

    '                Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
    '                Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)

    '                If ret = DialogResult.OK Then
    '                    FileName = OpenFile.FileName
    '                    OpenFile.Dispose()
    '                Else
    '                    System.Windows.Forms.Application.ExitThread()
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        objApplication.StatusBar.SetText(ex.Message)
    '        FileName = ""
    '    Finally
    '        OpenFile.Dispose()
    '    End Try

    'End Sub

    Public Sub New()
        MyBase.New()
        'Class_Initialize_Renamed()
        If Not Prj_AN.clsConnection.objApplication Is Nothing And Not Prj_AN.clsConnection.objCompany Is Nothing Then
            objApplication = Prj_AN.clsConnection.objApplication
            oCompany = Prj_AN.clsConnection.objCompany

        Else
            Dim oclsConnection = New clsConnection
            objApplication = Prj_AN.clsConnection.objApplication
            oCompany = Prj_AN.clsConnection.objCompany
        End If

        'SetFilter()
        subCreateMenu()
        'subCreateTable()
    End Sub

    Private Sub subCreateMenu()
        On Error GoTo ErrorHandler

        Dim objMenuCreation As SAPbouiCOM.MenuCreationParams = Nothing
        Dim objMenuItem As SAPbouiCOM.MenuItem = Nothing
        Dim strMenu As String

        objMenuCreation = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

        '=== Kelompok menu : Inventory
        objMenuItem = objApplication.Menus.Item(Inventory_MenuId)

        strMenu = ProjectMaster_MenuDesc

        If Not objMenuItem.SubMenus.Exists(ProjectMaster_MenuId) Then
            objMenuCreation.Type = SAPbouiCOM.BoMenuType.mt_STRING
            With objMenuCreation
                .UniqueID = ProjectMaster_MenuId
                .String = ProjectMaster_MenuDesc
                .Enabled = True
                .Position = 1
            End With

            objMenuItem.SubMenus.AddEx(objMenuCreation)
        End If

        If Not objMenuItem.SubMenus.Exists(ProjectHarvest_MenuId) Then
            objMenuCreation.Type = SAPbouiCOM.BoMenuType.mt_STRING
            With objMenuCreation
                .UniqueID = ProjectHarvest_MenuId
                .String = ProjectHarvest_MenuDesc
                .Enabled = True
                .Position = 2
            End With

            objMenuItem.SubMenus.AddEx(objMenuCreation)
        End If

        GoTo Setnothing

ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail adding menu " & strMenu & "!~6.0001~", vbExclamation, "SAP BO")
        End If

Setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMenuCreation)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMenuItem)
        objMenuCreation = Nothing
        objMenuItem = Nothing

    End Sub

    Private Sub subProjectHarvestScrPaint()
        subScrPaint("ScrProjectHarvest.srf", ProjectHarvest_FormId, intFormCountProjectHarvest, oFormProjectHarvest)

        subProjectHarvestFirstLoad(True, oFormProjectHarvest)

    End Sub

    Private Sub subScrPaintListGR()
        subScrPaint("ListGRPO.srf", ListGR_FormId, intFormCountListGR, objFormListGR)

        subGRPOFirstLoad(True, objFormListGR)

    End Sub

    Private Sub subScrPaintListInvTransfer()
        subScrPaint("ListTransfer.srf", ListInvTransfer_FormId, intFormCountListInvTransfer, objFormListInvTransfer)

        subInvTransferFirstLoad(True, objFormListInvTransfer)

    End Sub

    Private Sub SubScrPaintBatchFG()
        subScrPaint("BatchFinishGood.srf", BatchFG_FormId, intFormCountBatchFG, oFormBatchFG)

        subBatchFGFirstLoad(True, oFormBatchFG)

    End Sub

    Private Sub subScrPaintBatchRM()
        subScrPaint("BatchRawMaterial.srf", BatchRM_FormId, intFormCountBatchRM, oFormBatchRM)

        subBatchRMFirstLoad(True, oFormBatchRM)

    End Sub

    Private Sub subScrPaintListGI()
        subScrPaint("ListGI.srf", ListGI_FormId, intFormCountListGI, objFormListGI)

        subGIFirstLoad(True, objFormListGI)

    End Sub

    Private Sub subProjectMasterScrPaint()
        subScrPaint("ProjectMaster.srf", ProjectMaster_FormId, intFormCountProjectMaster, objFormProjectMaster)

        subProjectMasterSetFirstLoad(True, objFormProjectMaster)
    End Sub

    Private Sub subProjectHarvestFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)
        pForm.Freeze(True)
        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("Code", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            pForm.DataSources.UserDataSources.Add("MISPROID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 17)
            pForm.DataSources.UserDataSources.Add("MISNETID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            pForm.DataSources.UserDataSources.Add("MISSIGND", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISSCIES", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISESTSF", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISHARVP", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 4)
            pForm.DataSources.UserDataSources.Add("MISAGETR", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 4)
            pForm.DataSources.UserDataSources.Add("MISESTHD", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISESTLF", SAPbouiCOM.BoDataType.dt_PERCENT)
            pForm.DataSources.UserDataSources.Add("MISGENET", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("MISESTHQ", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISNFDIE", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISNETPUCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISGENCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            pForm.DataSources.UserDataSources.Add("MISPROSR", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            pForm.DataSources.UserDataSources.Add("MISHARVD", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISHARVQ", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISPROHR", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            pForm.DataSources.UserDataSources.Add("MISINIFQ", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISFEEDQ", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISFCR", SAPbouiCOM.BoDataType.dt_PERCENT)
            pForm.DataSources.UserDataSources.Add("MISFCE", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISTEFQK", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            pForm.DataSources.UserDataSources.Add("MISINIFC", SAPbouiCOM.BoDataType.dt_PRICE)
            pForm.DataSources.UserDataSources.Add("MISFEEDC", SAPbouiCOM.BoDataType.dt_PRICE)
            pForm.DataSources.UserDataSources.Add("MISTPCST", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISTPGRC", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISTPGRQ", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISPROCS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISNETST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)



        End If
        pForm.Items.Item("Code").Specific.databind.setbound(True, "", "Code")
        pForm.Items.Item("MISPROID").Specific.databind.setbound(True, "", "MISPROID")
        pForm.Items.Item("MISNETID").Specific.databind.setbound(True, "", "MISNETID")
        pForm.Items.Item("MISSIGND").Specific.databind.setbound(True, "", "MISSIGND")
        pForm.Items.Item("MISSCIES").Specific.databind.setbound(True, "", "MISSCIES")
        pForm.Items.Item("MISESTSF").Specific.databind.setbound(True, "", "MISESTSF")
        pForm.Items.Item("MISHARVP").Specific.databind.setbound(True, "", "MISHARVP")
        pForm.Items.Item("MISAGETR").Specific.databind.setbound(True, "", "MISAGETR")
        pForm.Items.Item("MISESTHD").Specific.databind.setbound(True, "", "MISESTHD")
        pForm.Items.Item("MISESTLF").Specific.databind.setbound(True, "", "MISESTLF")
        pForm.Items.Item("MISESTHQ").Specific.databind.setbound(True, "", "MISESTHQ")
        pForm.Items.Item("MISNFDIE").Specific.databind.setbound(True, "", "MISNFDIE")
        pForm.Items.Item("MISNETPUCD").Specific.databind.setbound(True, "", "MISNETPUCD")
        pForm.Items.Item("MISGENCD").Specific.databind.setbound(True, "", "MISGENCD")
        pForm.Items.Item("MISGENET").Specific.databind.setbound(True, "", "MISGENET")
        pForm.Items.Item("MISPROSR").Specific.databind.setbound(True, "", "MISPROSR")
        pForm.Items.Item("MISHARVD").Specific.databind.setbound(True, "", "MISHARVD")
        pForm.Items.Item("MISHARVQ").Specific.databind.setbound(True, "", "MISHARVQ")
        pForm.Items.Item("MISPROHR").Specific.databind.setbound(True, "", "MISPROHR")
        pForm.Items.Item("MISINIFQ").Specific.databind.setbound(True, "", "MISINIFQ")
        pForm.Items.Item("MISFEEDQ").Specific.databind.setbound(True, "", "MISFEEDQ")
        pForm.Items.Item("MISFCR").Specific.databind.setbound(True, "", "MISFCR")
        pForm.Items.Item("MISFCE").Specific.databind.setbound(True, "", "MISFCE")
        pForm.Items.Item("MISTEFQK").Specific.databind.setbound(True, "", "MISTEFQK")
        pForm.Items.Item("MISINIFC").Specific.databind.setbound(True, "", "MISINIFC")
        pForm.Items.Item("MISFEEDC").Specific.databind.setbound(True, "", "MISFEEDC")
        pForm.Items.Item("MISTPCST").Specific.databind.setbound(True, "", "MISTPCST")
        pForm.Items.Item("MISTPGRC").Specific.databind.setbound(True, "", "MISTPGRC")
        pForm.Items.Item("MISTPGRQ").Specific.databind.setbound(True, "", "MISTPGRQ")
        pForm.Items.Item("MISPROCS").Specific.databind.setbound(True, "", "MISPROCS")
        pForm.Items.Item("MISNETST").Specific.databind.setbound(True, "", "MISNETST")

        pForm.Items.Item("BtnNETID").Specific.Type = SAPbouiCOM.BoButtonTypes.bt_Image
        pForm.Items.Item("BtnNETID").Specific.Image = System.Windows.Forms.Application.StartupPath & "\" & pictCFL

        SubSetToolbar(oFormProjectHarvest, True, True, True, False, True, True, True, True, True, True, True, True, True, True, True)

        pForm.Freeze(False)

    End Sub

    Private Sub subProjectMasterSetFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)

        pForm.Freeze(True)

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("Code", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            pForm.DataSources.UserDataSources.Add("MISPROID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 17)
            pForm.DataSources.UserDataSources.Add("MISNETID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            pForm.DataSources.UserDataSources.Add("MISSIGND", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISSCIES", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISESTSF", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISHARVP", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 4)
            pForm.DataSources.UserDataSources.Add("MISAGETR", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 4)
            pForm.DataSources.UserDataSources.Add("MISESTHD", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISESTLF", SAPbouiCOM.BoDataType.dt_PERCENT)
            pForm.DataSources.UserDataSources.Add("MISGENET", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("MISESTHQ", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISNFDIE", SAPbouiCOM.BoDataType.dt_QUANTITY, 9)
            pForm.DataSources.UserDataSources.Add("MISNETPUCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISGENCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            pForm.DataSources.UserDataSources.Add("MISPROSR", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            pForm.DataSources.UserDataSources.Add("MISHARVD", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("MISHARVQ", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISPROHR", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            pForm.DataSources.UserDataSources.Add("MISINIFQ", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISFEEDQ", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISFCR", SAPbouiCOM.BoDataType.dt_PERCENT)
            pForm.DataSources.UserDataSources.Add("MISFCE", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISTEFQK", SAPbouiCOM.BoDataType.dt_QUANTITY, 10)
            pForm.DataSources.UserDataSources.Add("MISINIFC", SAPbouiCOM.BoDataType.dt_PRICE)
            pForm.DataSources.UserDataSources.Add("MISFEEDC", SAPbouiCOM.BoDataType.dt_PRICE)
            pForm.DataSources.UserDataSources.Add("MISTPCST", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISTPGRC", SAPbouiCOM.BoDataType.dt_SUM)
            pForm.DataSources.UserDataSources.Add("MISTPGRQ", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("MISPROCS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISNETST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("MISHATGO", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 20)



        End If
        pForm.Items.Item("Code").Specific.databind.setbound(True, "", "Code")
        pForm.Items.Item("MISPROID").Specific.databind.setbound(True, "", "MISPROID")
        pForm.Items.Item("MISNETID").Specific.databind.setbound(True, "", "MISNETID")
        pForm.Items.Item("MISSIGND").Specific.databind.setbound(True, "", "MISSIGND")
        pForm.Items.Item("MISSCIES").Specific.databind.setbound(True, "", "MISSCIES")
        pForm.Items.Item("MISESTSF").Specific.databind.setbound(True, "", "MISESTSF")
        pForm.Items.Item("MISHARVP").Specific.databind.setbound(True, "", "MISHARVP")
        pForm.Items.Item("MISAGETR").Specific.databind.setbound(True, "", "MISAGETR")
        pForm.Items.Item("MISESTHD").Specific.databind.setbound(True, "", "MISESTHD")
        pForm.Items.Item("MISESTLF").Specific.databind.setbound(True, "", "MISESTLF")
        pForm.Items.Item("MISESTHQ").Specific.databind.setbound(True, "", "MISESTHQ")
        pForm.Items.Item("MISNFDIE").Specific.databind.setbound(True, "", "MISNFDIE")
        pForm.Items.Item("MISNETPUCD").Specific.databind.setbound(True, "", "MISNETPUCD")
        pForm.Items.Item("MISGENCD").Specific.databind.setbound(True, "", "MISGENCD")
        pForm.Items.Item("MISGENET").Specific.databind.setbound(True, "", "MISGENET")
        pForm.Items.Item("MISPROSR").Specific.databind.setbound(True, "", "MISPROSR")
        pForm.Items.Item("MISHARVD").Specific.databind.setbound(True, "", "MISHARVD")
        pForm.Items.Item("MISHARVQ").Specific.databind.setbound(True, "", "MISHARVQ")
        pForm.Items.Item("MISPROHR").Specific.databind.setbound(True, "", "MISPROHR")
        pForm.Items.Item("MISINIFQ").Specific.databind.setbound(True, "", "MISINIFQ")
        pForm.Items.Item("MISFEEDQ").Specific.databind.setbound(True, "", "MISFEEDQ")
        pForm.Items.Item("MISINIFC").Specific.databind.setbound(True, "", "MISINIFC")
        pForm.Items.Item("MISFEEDC").Specific.databind.setbound(True, "", "MISFEEDC")
        pForm.Items.Item("MISFCR").Specific.databind.setbound(True, "", "MISFCR")
        pForm.Items.Item("MISFCE").Specific.databind.setbound(True, "", "MISFCE")
        pForm.Items.Item("MISTEFQK").Specific.databind.setbound(True, "", "MISTEFQK")
        pForm.Items.Item("MISTPCST").Specific.databind.setbound(True, "", "MISTPCST")
        pForm.Items.Item("MISTPGRC").Specific.databind.setbound(True, "", "MISTPGRC")
        pForm.Items.Item("MISTPGRQ").Specific.databind.setbound(True, "", "MISTPGRQ")
        pForm.Items.Item("MISPROCS").Specific.databind.setbound(True, "", "MISPROCS")
        pForm.Items.Item("MISNETST").Specific.databind.setbound(True, "", "MISNETST")
        pForm.Items.Item("MISHATGO").Specific.databind.setbound(True, "", "MISHATGO")

        pForm.Items.Item("BtnNETID").Specific.Type = SAPbouiCOM.BoButtonTypes.bt_Image
        pForm.Items.Item("BtnNETID").Specific.Image = System.Windows.Forms.Application.StartupPath & "\" & pictCFL

        'pForm.Items.Item("BtnSCIES").Specific.Type = SAPbouiCOM.BoButtonTypes.bt_Image
        'pForm.Items.Item("BtnSCIES").Specific.Image = Application.StartupPath & "\" & pictCFL

        pForm.Items.Item("BtnGENCD").Specific.Type = SAPbouiCOM.BoButtonTypes.bt_Image
        pForm.Items.Item("BtnGENCD").Specific.Image = System.Windows.Forms.Application.StartupPath & "\" & pictCFL
        'pForm.Items.Item("CFLNETID").Specific.picture = Application.StartupPath & "\" & "CFL"
        'SubSearchDataProjectMaster(pForm)



        SubSetToolbar(objFormProjectMaster, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)

        pForm.Freeze(False)


    End Sub

    Private Sub subBatchRMFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("NetCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            pForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("RitNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2)
            pForm.DataSources.UserDataSources.Add("BoxNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2)
        End If

        pForm.Items.Item("NetCd").Specific.databind.setbound(True, "", "NetCd")
        pForm.Items.Item("DocDate").Specific.databind.setbound(True, "", "DocDate")
        pForm.Items.Item("RitNo").Specific.databind.setbound(True, "", "RitNo")
        pForm.Items.Item("BoxNo").Specific.databind.setbound(True, "", "BoxNo")

    End Sub

    Private Sub subBatchFGFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("Region", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            pForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 7)
        End If

        pForm.Items.Item("Region").Specific.databind.setbound(True, "", "Region")
        pForm.Items.Item("DocDate").Specific.databind.setbound(True, "", "DocDate")
        pForm.Items.Item("No").Specific.databind.setbound(True, "", "No")

    End Sub

    Private Sub subGIFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        oMatrix = pForm.Items.Item("MtxGI").Specific
        oColumns = oMatrix.Columns

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Number", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Series", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            pForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("RitNo", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)

            pForm.DataSources.UserDataSources.Add("NetId", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 17)
            pForm.DataSources.UserDataSources.Add("Ref", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            pForm.DataSources.UserDataSources.Add("Remarks", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            pForm.DataSources.UserDataSources.Add("QuantityKg", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("QuantityPc", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 11)

        End If

        oColumns.Item("No").DataBind.SetBound(True, "", "No")
        oColumns.Item("DocEntry").DataBind.SetBound(True, "", "DocEntry")
        oColumns.Item("Number").DataBind.SetBound(True, "", "Number")
        oColumns.Item("Series").DataBind.SetBound(True, "", "Series")
        oColumns.Item("DocDate").DataBind.SetBound(True, "", "DocDate")
        oColumns.Item("RitNo").DataBind.SetBound(True, "", "RitNo")

        oColumns.Item("NetId").DataBind.SetBound(True, "", "NetId")
        oColumns.Item("Ref").DataBind.SetBound(True, "", "Ref")
        oColumns.Item("Remarks").DataBind.SetBound(True, "", "Remarks")
        oColumns.Item("QuantityKg").DataBind.SetBound(True, "", "QuantityKg")
        oColumns.Item("QuantityPc").DataBind.SetBound(True, "", "QuantityPc")

        oColumns.Item("DocEntry").Editable = False
        oColumns.Item("Number").Editable = False
        oColumns.Item("Series").Editable = False
        oColumns.Item("DocDate").Editable = False
        oColumns.Item("RitNo").Editable = False
        oColumns.Item("NetId").Editable = False
        oColumns.Item("Ref").Editable = False
        oColumns.Item("Remarks").Editable = False
        oColumns.Item("QuantityKg").Editable = False
        oColumns.Item("QuantityPc").Editable = False
        'End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)

        SubSearchDataGI(pForm)

    End Sub

    Private Sub subGRPOFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        oMatrix = pForm.Items.Item("MtxGRPO").Specific
        oColumns = oMatrix.Columns

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Number", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Series", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            pForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("RitNo", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)

            pForm.DataSources.UserDataSources.Add("NetId", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 17)
            pForm.DataSources.UserDataSources.Add("Ref", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            pForm.DataSources.UserDataSources.Add("Remarks", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            pForm.DataSources.UserDataSources.Add("QuantityKg", SAPbouiCOM.BoDataType.dt_QUANTITY)
            pForm.DataSources.UserDataSources.Add("QuantityPc", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 11)

        End If

        oColumns.Item("No").DataBind.SetBound(True, "", "No")
        oColumns.Item("DocEntry").DataBind.SetBound(True, "", "DocEntry")
        oColumns.Item("Number").DataBind.SetBound(True, "", "Number")
        oColumns.Item("Series").DataBind.SetBound(True, "", "Series")
        oColumns.Item("DocDate").DataBind.SetBound(True, "", "DocDate")
        oColumns.Item("RitNo").DataBind.SetBound(True, "", "RitNo")

        oColumns.Item("NetId").DataBind.SetBound(True, "", "NetId")
        oColumns.Item("Ref").DataBind.SetBound(True, "", "Ref")
        oColumns.Item("Remarks").DataBind.SetBound(True, "", "Remarks")
        oColumns.Item("QuantityKg").DataBind.SetBound(True, "", "QuantityKg")
        oColumns.Item("QuantityPc").DataBind.SetBound(True, "", "QuantityPc")

        oColumns.Item("DocEntry").Editable = False
        oColumns.Item("Number").Editable = False
        oColumns.Item("Series").Editable = False
        oColumns.Item("DocDate").Editable = False
        oColumns.Item("RitNo").Editable = False
        oColumns.Item("NetId").Editable = False
        oColumns.Item("Ref").Editable = False
        oColumns.Item("Remarks").Editable = False
        oColumns.Item("QuantityKg").Editable = False
        oColumns.Item("QuantityPc").Editable = False
        'End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)

        SubSearchDataGRPO(pForm)

    End Sub

    Private Sub subInvTransferFirstLoad(ByVal pFirstLoad As Boolean, ByVal pForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        oMatrix = pForm.Items.Item("MtxTrnsfr").Specific
        oColumns = oMatrix.Columns

        If pFirstLoad Then
            pForm.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Number", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            pForm.DataSources.UserDataSources.Add("Series", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            pForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            pForm.DataSources.UserDataSources.Add("FrmWhs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("ReqUsr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            pForm.DataSources.UserDataSources.Add("RowNo", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            pForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 30)


        End If

        oColumns.Item("No").DataBind.SetBound(True, "", "No")
        oColumns.Item("DocEntry").DataBind.SetBound(True, "", "DocEntry")
        oColumns.Item("Number").DataBind.SetBound(True, "", "Number")
        oColumns.Item("Series").DataBind.SetBound(True, "", "Series")
        oColumns.Item("DocDate").DataBind.SetBound(True, "", "DocDate")
        oColumns.Item("FrmWhs").DataBind.SetBound(True, "", "FrmWhs")
        oColumns.Item("ReqUsr").DataBind.SetBound(True, "", "ReqUsr")
        oColumns.Item("RowNo").DataBind.SetBound(True, "", "RowNo")
        oColumns.Item("ItemCode").DataBind.SetBound(True, "", "ItemCode")

        oColumns.Item("DocEntry").Editable = False
        oColumns.Item("Number").Editable = False
        oColumns.Item("Series").Editable = False
        oColumns.Item("DocDate").Editable = False
        oColumns.Item("FrmWhs").Editable = False
        oColumns.Item("ReqUsr").Editable = False
        oColumns.Item("RowNo").Editable = False
        oColumns.Item("ItemCode").Editable = False

        'End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)

        SubSearchDataInvTransfer(pForm)

    End Sub

    Private Sub SubSearchDataGI(ByVal pForm As SAPbouiCOM.Form)
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim StrSql As String
        Dim intLoop As Integer
        Dim oUserDataSource(11) As SAPbouiCOM.UserDataSource
        Dim TransactionType As Integer
        Dim Warehouse As String
        'Dim oLinkColumn As SAPbouiCOM.LinkedButton

        objMatrix = objFormListGI.Items.Item("MtxGI").Specific
        objColumns = objMatrix.Columns

        oUserDataSource(0) = objFormListGI.DataSources.UserDataSources.Item("No")
        oUserDataSource(1) = objFormListGI.DataSources.UserDataSources.Item("DocEntry")
        oUserDataSource(2) = objFormListGI.DataSources.UserDataSources.Item("Number")
        oUserDataSource(3) = objFormListGI.DataSources.UserDataSources.Item("Series")
        oUserDataSource(4) = objFormListGI.DataSources.UserDataSources.Item("DocDate")
        oUserDataSource(5) = objFormListGI.DataSources.UserDataSources.Item("RitNo")

        oUserDataSource(6) = objFormListGI.DataSources.UserDataSources.Item("NetId")
        oUserDataSource(7) = objFormListGI.DataSources.UserDataSources.Item("Ref")
        oUserDataSource(8) = objFormListGI.DataSources.UserDataSources.Item("Remarks")
        oUserDataSource(9) = objFormListGI.DataSources.UserDataSources.Item("QuantityKg")
        oUserDataSource(10) = objFormListGI.DataSources.UserDataSources.Item("QuantityPc")

        objMatrix.Clear()

        objFormListGI.Title = "ListGI"

        TransactionType = objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string



        If TransactionType = 2 Then
            If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                objFormListGI.Close()
                objApplication.StatusBar.SetText("You Must Input Transaction Warehouse ~20.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            Else
                Warehouse = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string

                StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                         "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                         "FROM OIGE T0 " & _
                         "INNER JOIN IGE1 T1 " & _
                         "ON T0.DocEntry = T1.DocEntry " & _
                        "WHERE T0.U_MISTRXTP = 3 AND T0.U_MISDESTW = '" & Warehouse & "' AND NOT EXISTS (SELECT G1.DocNum FROM OIGN G1 WHERE G1.U_MISTRXTP = 2 AND T0.DocNum = G1.U_MISREFFD) " & _
                        " AND NOT EXISTS(SELECT G1.Docnum FROM OIGN G1 WHERE G1.U_MISTRXTP = 5 AND T0.DocNum = G1.U_MISREFFD) " & _
                        "GROUP BY T0.DocEntry, T0.DocNum, T0.Series, T0.DocDate, U_MISRITNO, T1.U_MISNETID, T0.Ref2, T0.Comments  "
                '"INNER JOIN [@MIS_PRJMSTR] T2 " & _
                '"ON T1.U_MISNETID = T2.U_MISNETID " & _
                '"OR T1.DocDate = T2.U_MISSIGND " & _
                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            End If
        ElseIf TransactionType = 3 Then
            If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                objFormListGI.Close()
                objApplication.StatusBar.SetText("You Must Input Transaction Warehouse ~20.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            Else
                Warehouse = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string

                StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                         "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                         "FROM OIGE T0 " & _
                         "INNER JOIN IGE1 T1 " & _
                         "ON T0.DocEntry = T1.DocEntry " & _
                         "WHERE T0.U_MISTRXTP = 5 AND T0.U_MISDESTW = '" & Warehouse & "' AND NOT EXISTS (SELECT G1.DocNum FROM OIGN G1 WHERE G1.U_MISTRXTP = 3 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "AND NOT EXISTS(SELECT G1.Docnum FROM OIGN G1 WHERE G1.U_MISTRXTP = 5 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "GROUP BY T0.DocEntry, T0.DocNum, T0.Series, T0.DocDate, U_MISRITNO, T1.U_MISNETID, T0.Ref2, T0.Comments  "
                '"INNER JOIN [@MIS_PRJMSTR] T2 " & _
                '"ON T1.U_MISNETID = T2.U_MISNETID " & _
                '"OR T1.DocDate = T2.U_MISSIGND " & _
                '"WHERE T2.U_MISNETST = 'H' AND T0.U_MISTRXTP = 5 "
                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            End If

        ElseIf TransactionType = 4 Then
            If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                objFormListGI.Close()
                objApplication.StatusBar.SetText("You Must Input Transaction Warehouse ~20.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            Else
                Warehouse = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string
                StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                         "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                         "FROM OIGE T0 " & _
                         "INNER JOIN IGE1 T1 " & _
                         "ON T0.DocEntry = T1.DocEntry " & _
                         "WHERE T0.U_MISTRXTP = 6 AND LEFT(T1.WhsCode,1) = LEFT('" & Warehouse & "',1) AND NOT EXISTS (SELECT G1.DocNum FROM OIGN G1 WHERE G1.U_MISTRXTP = 4 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "AND NOT EXISTS(SELECT G1.Docnum FROM OIGN G1 WHERE G1.U_MISTRXTP = 5 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "GROUP BY T0.DocEntry, T0.DocNum, T0.Series, T0.DocDate, U_MISRITNO, T1.U_MISNETID, T0.Ref2, T0.Comments  "
                '"INNER JOIN [@MIS_PRJMSTR] T2 " & _
                '"ON T1.U_MISNETID = T2.U_MISNETID " & _
                '"OR T1.DocDate = T2.U_MISSIGND " & _
                '"WHERE T2.U_MISNETST = 'H' AND T0.U_MISTRXTP = 5 "
                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            End If
        Else
            Warehouse = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string
            StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                     "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                     "FROM OIGE T0 " & _
                     "INNER JOIN IGE1 T1 " & _
                     "ON T0.DocEntry = T1.DocEntry " & _
                     "WHERE T0.U_MISDESTW = '" & Warehouse & "' AND NOT EXISTS (SELECT G1.DocNum FROM OIGN G1 WHERE T0.DocNum = G1.U_MISREFFD) " & _
                     "GROUP BY T0.DocEntry, T0.DocNum, T0.Series, T0.DocDate, U_MISRITNO, T1.U_MISNETID, T0.Ref2, T0.Comments  "
            '"INNER JOIN [@MIS_PRJMSTR] T2 " & _
            '"ON T1.U_MISNETID = T2.U_MISNETID " & _
            '"OR T1.DocDate = T2.U_MISSIGND " & _
            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecSet.DoQuery(StrSql)

            End If



            If objRecSet.RecordCount > 0 Then
                intLoop = 0
                Do While Not objRecSet.EoF
                    intLoop = intLoop + 1
                    oUserDataSource(0).Value = intLoop
                    oUserDataSource(1).Value = objRecSet.Fields.Item("DocEntry").Value
                    oUserDataSource(2).Value = objRecSet.Fields.Item("Number").Value
                    oUserDataSource(3).Value = objRecSet.Fields.Item("Series").Value
                    oUserDataSource(4).Value = ClsGlobal.fctFormatDate(objRecSet.Fields.Item("DocDate").Value, oCompany)
                    oUserDataSource(5).Value = objRecSet.Fields.Item("RitNo").Value
                    oUserDataSource(6).Value = objRecSet.Fields.Item("NetId").Value
                    oUserDataSource(7).Value = objRecSet.Fields.Item("Ref").Value
                    oUserDataSource(8).Value = objRecSet.Fields.Item("Remarks").Value
                    oUserDataSource(9).Value = objRecSet.Fields.Item("QuantityKg").Value
                    oUserDataSource(10).Value = objRecSet.Fields.Item("QuantityPc").Value
                    objMatrix.AddRow()
                    objRecSet.MoveNext()
                Loop

                'oLinkColumn = objColumns.Item("DocEntry").Cells.Item(0).Specific.caption
                'oLinkColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_GoodsIssue
                'oLinkColumn.Editable = False

                SubSetToolbar(objFormListGI, False, False, False, False, False, False, False, False, _
                                            False, False, False, False, True, True, True)

            Else
                objApplication.StatusBar.SetText("No Data ~20.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objFormListGI.Close()
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)

Setnothing:

            objRecSet = Nothing
            objMatrix = Nothing
            objColumns = Nothing
            oUserDataSource(0) = Nothing
            oUserDataSource(1) = Nothing
            oUserDataSource(2) = Nothing
            oUserDataSource(3) = Nothing
            oUserDataSource(4) = Nothing
            oUserDataSource(5) = Nothing
            oUserDataSource(6) = Nothing
            oUserDataSource(7) = Nothing
            oUserDataSource(8) = Nothing
            oUserDataSource(9) = Nothing
            oUserDataSource(10) = Nothing
    End Sub

    Private Sub SubSearchDataGRPO(ByVal pForm As SAPbouiCOM.Form)
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim StrSql As String
        Dim intLoop As Integer
        Dim oUserDataSource(11) As SAPbouiCOM.UserDataSource
        Dim TransactionType As Integer
        Dim TanggalHarvest As Date
        Dim Warehouse As String

        TanggalHarvest = ClsGlobal.fctFormatDateSave(oCompany, objFormGoodIssue.Items.Item("9").Specific.string, 1)
        TransactionType = objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string

        objMatrix = objFormListGR.Items.Item("MtxGRPO").Specific
        objColumns = objMatrix.Columns

        oUserDataSource(0) = objFormListGR.DataSources.UserDataSources.Item("No")
        oUserDataSource(1) = objFormListGR.DataSources.UserDataSources.Item("DocEntry")
        oUserDataSource(2) = objFormListGR.DataSources.UserDataSources.Item("Number")
        oUserDataSource(3) = objFormListGR.DataSources.UserDataSources.Item("Series")
        oUserDataSource(4) = objFormListGR.DataSources.UserDataSources.Item("DocDate")
        oUserDataSource(5) = objFormListGR.DataSources.UserDataSources.Item("RitNo")

        oUserDataSource(6) = objFormListGR.DataSources.UserDataSources.Item("NetId")
        oUserDataSource(7) = objFormListGR.DataSources.UserDataSources.Item("Ref")
        oUserDataSource(8) = objFormListGR.DataSources.UserDataSources.Item("Remarks")
        oUserDataSource(9) = objFormListGR.DataSources.UserDataSources.Item("QuantityKg")
        oUserDataSource(10) = objFormListGR.DataSources.UserDataSources.Item("QuantityPc")

        objMatrix.Clear()

        objFormListGR.Title = "ListGRPO"

        If TransactionType = 5 Then
            If objFormGoodIssueUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                objFormListGR.Close()
                objApplication.MessageBox("Must Fill transaction Warehouse", 1, "OK")
                GoTo Setnothing
            Else

                Warehouse = objFormGoodIssueUDF.Items.Item("U_MISTRXWH").Specific.string

                StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                         "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                         "FROM OIGN T0 " & _
                         "INNER JOIN IGN1 T1 " & _
                         "ON T0.DocEntry = T1.DocEntry " & _
                         "INNER JOIN [@MIS_PRJMSTR] T2 " & _
                         "ON T1.U_MISPROID = T2.U_MISPROID LEFT JOIN OITM T3 ON T1.ItemCode = T3.ItemCode INNER JOIN OITB T4 ON T3.ItmsGrpCod = T4.ItmsGrpCod " & _
                         "WHERE T0.U_MISTRXTP = 1 AND LEFT(T1.WhsCode,1) = LEFT('" & Warehouse & "',1) AND T4.ItmsGrpNam = 'RM Biomass'  AND T0.DocDate = '" & TanggalHarvest & "' AND NOT EXISTS (SELECT G1.DocNum FROM OIGE G1 WHERE G1.U_MISTRXTP = 5 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "AND NOT EXISTS (SELECT G1.DocNum FROM OIGE G1 WHERE G1.U_MISTRXTP = 8 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "GROUP BY T0.DocEntry, T0.DocNum , T0.Series , T0.DocDate , U_MISRITNO, U_MISRITNO, T1.U_MISNETID , T0.Ref2, T0.Comments"
                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            End If
        ElseIf TransactionType = 6 Then
            If objFormGoodIssueUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                objFormListGR.Close()
                objApplication.MessageBox("Must Fill transaction Warehouse", 1, "OK")
                GoTo Setnothing
            Else

                Warehouse = objFormGoodIssueUDF.Items.Item("U_MISTRXWH").Specific.string
                StrSql = " SELECT DISTINCT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, " & _
                         "T0.DocDate DocDate, ISNULL(U_MISRITNO,0) RitNo, T1.U_MISNETID NetId, T0.Ref2 Ref, T0.Comments Remarks, SUM(T1.quantity) QuantityKg, SUM(T1.U_MISFISHQ) QuantityPc " & _
                         "FROM OIGN T0 " & _
                         "INNER JOIN IGN1 T1 " & _
                         "ON T0.DocEntry = T1.DocEntry " & _
                         "WHERE T0.U_MISTRXTP = 3 AND LEFT(T1.WhsCode,1) = LEFT('" & Warehouse & "',1) AND NOT EXISTS (SELECT G1.DocNum FROM OIGE G1 WHERE G1.U_MISTRXTP = 6 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "AND NOT EXISTS (SELECT G1.DocNum FROM OIGE G1 WHERE G1.U_MISTRXTP = 7 AND T0.DocNum = G1.U_MISREFFD) " & _
                         "GROUP BY T0.DocEntry, T0.DocNum , T0.Series , T0.DocDate , U_MISRITNO, U_MISRITNO, T1.U_MISNETID , T0.Ref2, T0.Comments"
                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            End If
        Else
                objFormListGR.Close()
                objApplication.StatusBar.SetText("Transaction Type Not Found ~21.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objRecSet.RecordCount > 0 Then
                intLoop = 0
                Do While Not objRecSet.EoF
                    intLoop = intLoop + 1
                    oUserDataSource(0).Value = intLoop
                    oUserDataSource(1).Value = objRecSet.Fields.Item("DocEntry").Value
                    oUserDataSource(2).Value = objRecSet.Fields.Item("Number").Value
                    oUserDataSource(3).Value = objRecSet.Fields.Item("Series").Value
                    oUserDataSource(4).Value = ClsGlobal.fctFormatDate(objRecSet.Fields.Item("DocDate").Value, oCompany)
                    oUserDataSource(5).Value = objRecSet.Fields.Item("RitNo").Value
                    oUserDataSource(6).Value = objRecSet.Fields.Item("NetId").Value
                    oUserDataSource(7).Value = objRecSet.Fields.Item("Ref").Value
                    oUserDataSource(8).Value = objRecSet.Fields.Item("Remarks").Value
                    oUserDataSource(9).Value = objRecSet.Fields.Item("QuantityKg").Value
                    oUserDataSource(10).Value = objRecSet.Fields.Item("QuantityPc").Value
                    objMatrix.AddRow()
                    objRecSet.MoveNext()
                Loop

                SubSetToolbar(objFormListGR, False, False, False, False, False, False, False, False, _
                                            False, False, False, False, True, True, True)

            Else
                objApplication.StatusBar.SetText("No Data ~21.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objFormListGR.Close()
            End If



            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
            objRecSet = Nothing
            objMatrix = Nothing
            objColumns = Nothing
            oUserDataSource(0) = Nothing
            oUserDataSource(1) = Nothing
            oUserDataSource(2) = Nothing
            oUserDataSource(3) = Nothing
            oUserDataSource(4) = Nothing
            oUserDataSource(5) = Nothing
            oUserDataSource(6) = Nothing
            oUserDataSource(7) = Nothing
            oUserDataSource(8) = Nothing
            oUserDataSource(9) = Nothing
            oUserDataSource(10) = Nothing

Setnothing:
            TransactionType = Nothing
    End Sub

    Private Sub SubSearchDataInvTransfer(ByVal pForm As SAPbouiCOM.Form)
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim StrSql As String
        Dim intLoop As Integer
        Dim oUserDataSource(9) As SAPbouiCOM.UserDataSource
        Dim TransactionType As Integer

        objMatrix = objFormListInvTransfer.Items.Item("MtxTrnsfr").Specific
        objColumns = objMatrix.Columns

        oUserDataSource(0) = objFormListInvTransfer.DataSources.UserDataSources.Item("No")
        oUserDataSource(1) = objFormListInvTransfer.DataSources.UserDataSources.Item("DocEntry")
        oUserDataSource(2) = objFormListInvTransfer.DataSources.UserDataSources.Item("Number")
        oUserDataSource(3) = objFormListInvTransfer.DataSources.UserDataSources.Item("Series")
        oUserDataSource(4) = objFormListInvTransfer.DataSources.UserDataSources.Item("DocDate")
        oUserDataSource(5) = objFormListInvTransfer.DataSources.UserDataSources.Item("FrmWhs")
        oUserDataSource(6) = objFormListInvTransfer.DataSources.UserDataSources.Item("ReqUsr")
        oUserDataSource(7) = objFormListInvTransfer.DataSources.UserDataSources.Item("RowNo")
        oUserDataSource(8) = objFormListInvTransfer.DataSources.UserDataSources.Item("ItemCode")

        objMatrix.Clear()

        objFormListInvTransfer.Title = "ListTransfer"
        If objFormInvTransferUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Choose Transaction Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objFormListInvTransfer.Close()
            GoTo Setnothing
        Else
            If objFormInvTransferUDF.Items.Item("U_MISTRXTP").Specific.string = 2 Then

                StrSql = "SELECT T0.DocEntry DocEntry, T0.DocNum Number, T0.Series Series, T0.DocDate DocDate, T0.Filler FrmWhs, T0.U_MISREQBY ReqUsr, T1.VisOrder RowNo, T1.ItemCode ItemCode FROM OWTR T0 " & _
                        "Inner Join WTR1 T1 " & _
                        "ON T0.DocEntry = T1. DocEntry " & _
                        "WHERE(U_MISTRXTP = 1) " & _
                        "AND NOT EXISTS( " & _
                        "SELECT T3.Filler FromWhsCode, T4.WhsCode,  T3.U_MISPRNO, T4.ItemCode , T4.Quantity FROM OWTR T3 " & _
                        "Inner Join WTR1 T4 " & _
                        "ON T3.DocEntry = T4. DocEntry " & _
                        "WHERE(U_MISTRXTP = 2) " & _
                        "AND T1.WhsCode = T3.Filler " & _
                        "AND T1.ItemCode = T4.ItemCode " & _
                        "AND T0.U_MISPRNO = T3.U_MISPRNO " & _
                        "AND T1.Quantity = T4.Quantity) "

                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(StrSql)
            Else
                objApplication.StatusBar.SetText("No transaction ~22.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objFormListInvTransfer.Close()
                GoTo Setnothing
            End If

        End If


        If objRecSet.RecordCount > 0 Then
            intLoop = 0
            Do While Not objRecSet.EoF
                intLoop = intLoop + 1
                oUserDataSource(0).Value = intLoop
                oUserDataSource(1).Value = objRecSet.Fields.Item("DocEntry").Value
                oUserDataSource(2).Value = objRecSet.Fields.Item("Number").Value
                oUserDataSource(3).Value = objRecSet.Fields.Item("Series").Value
                oUserDataSource(4).Value = ClsGlobal.fctFormatDate(objRecSet.Fields.Item("DocDate").Value, oCompany)
                oUserDataSource(5).Value = objRecSet.Fields.Item("FrmWhs").Value
                oUserDataSource(6).Value = objRecSet.Fields.Item("ReqUsr").Value
                oUserDataSource(7).Value = objRecSet.Fields.Item("RowNo").Value
                oUserDataSource(8).Value = objRecSet.Fields.Item("ItemCode").Value
                objMatrix.AddRow()
                objRecSet.MoveNext()
            Loop

            SubSetToolbar(objFormListInvTransfer, False, False, False, False, False, False, False, False, _
                                        False, False, False, False, True, True, True)

        Else
            objFormListInvTransfer.Close()
            objApplication.StatusBar.SetText("No Data ~22.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
        objRecSet = Nothing
        objMatrix = Nothing
        objColumns = Nothing
        oUserDataSource(0) = Nothing
        oUserDataSource(1) = Nothing
        oUserDataSource(2) = Nothing
        oUserDataSource(3) = Nothing
        oUserDataSource(4) = Nothing
        oUserDataSource(5) = Nothing
        oUserDataSource(6) = Nothing
        oUserDataSource(7) = Nothing
        oUserDataSource(8) = Nothing

Setnothing:
        TransactionType = Nothing
    End Sub

    Private Sub SubSetToolbar(ByVal pForm As SAPbouiCOM.Form, _
    Optional ByVal pblnFirstLoad As Boolean = True, _
    Optional ByVal pblnLockScr As Boolean = False, _
    Optional ByVal pblnFind As Boolean = False, _
    Optional ByVal pblnAdd As Boolean = False, _
    Optional ByVal pblnFirst As Boolean = False, _
    Optional ByVal pblnNext As Boolean = False, _
    Optional ByVal pblnPrev As Boolean = False, _
    Optional ByVal pblnLast As Boolean = False, _
    Optional ByVal pblnFilterTbl As Boolean = False, _
    Optional ByVal pblnClose As Boolean = False, _
    Optional ByVal pblnDuplicate As Boolean = False, _
    Optional ByVal pblnCloseRow As Boolean = False, _
    Optional ByVal pblnFSetting As Boolean = False, _
    Optional ByVal pblnQManager As Boolean = False, _
    Optional ByVal pblnPrintPrv As Boolean = False)

        pForm.EnableMenu(TBar_LockScr, pblnFind)
        pForm.EnableMenu(TBar_Find, pblnFind)
        pForm.EnableMenu(TBar_Add, pblnAdd)
        pForm.EnableMenu(TBar_First, pblnFirst)
        pForm.EnableMenu(TBar_Next, pblnNext)
        pForm.EnableMenu(TBar_Prev, pblnPrev)
        pForm.EnableMenu(TBar_Last, pblnLast)
        pForm.EnableMenu(TBar_FilterTbl, pblnFilterTbl)
        pForm.EnableMenu(TBar_Close, pblnClose)
        pForm.EnableMenu(TBar_Duplicate, pblnDuplicate)
        pForm.EnableMenu(TBar_CloseRow, pblnCloseRow)
        pForm.EnableMenu(TBar_FSetting, pblnFSetting)
        pForm.EnableMenu(TBar_QManager, pblnQManager)
        pForm.EnableMenu(TBar_PrintPrv, pblnPrintPrv)

    End Sub

    Private Function fctFormExist(ByVal pFormId As String, ByRef pLoop As Integer) As Boolean
        Dim objForms As SAPbouiCOM.Forms
        Dim intLoop As Integer

        fctFormExist = False
        pLoop = 0

        objForms = objApplication.Forms

        If objForms.Count > 0 Then
            For intLoop = 0 To objForms.Count - 1
                If objForms.Item(intLoop).Type = pFormId Then
                    fctFormExist = True
                    pLoop = intLoop
                    objForms = Nothing
                    Exit Function
                End If
            Next
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objForms)
        objForms = Nothing
    End Function

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent

        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                'If (EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown) Then

                System.Windows.Forms.Application.Exit()

                objApplication.MessageBox("A Shut Down Event has been caught" _
                    & vbNewLine & "Terminating Add On...")
                'End If
                End
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                'If (EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown) Then

                System.Windows.Forms.Application.Exit()

                objApplication.MessageBox("A Company Changed Event has been caught" _
                    & vbNewLine & "Terminating Add On...")
                'End If
                End
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition

                System.Windows.Forms.Application.Exit()

                objApplication.MessageBox("A Server Termination Event has been caught" _
                    & vbNewLine & "Terminating Add On...")
                'End If
                End

        End Select
    End Sub

    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent

        ' Inventory Transfer
        If pVal.FormTypeEx = InvTransfer_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormInvTransfer = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction Then
                        subInvTransferAddObject(objFormInvTransfer, True)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "btnTrnsfr"
                            If pVal.BeforeAction Then
                                If fctFormExist(ListInvTransfer_FormId, intFormCountListInvTransfer) Then
                                    objApplication.Forms.Item(intFormCountListInvTransfer).Select()
                                Else
                                    subScrPaintListInvTransfer()
                                End If

                            End If

                    End Select
            End Select
        End If

        ' List Inventory Transfer
        If pVal.FormTypeEx = ListInvTransfer_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormListInvTransfer = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    'If fctFormExist(GoodIssue_FormId, intRowGoodIssueDetail) Then
                    '    objApplication.Forms.Item(intRowGoodIssueDetail).Select()
                    'Else
                    Dim DocEntry As Integer
                    Dim RowNo As Integer
                    If (pVal.ColUID = "DocEntry" Or pVal.ColUID = "Number" Or pVal.ColUID = "series" Or pVal.ColUID = "DocDate" Or pVal.ColUID = "FrmWhs") And (pVal.BeforeAction) And (pVal.Row > 0) Then
                        blnModalListGr = False
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing

                        objMatrix = objFormListInvTransfer.Items.Item("MtxTrnsfr").Specific
                        objColumns = objMatrix.Columns
                        DocEntry = objColumns.Item("DocEntry").Cells.Item(pVal.Row).Specific.string
                        RowNo = objColumns.Item("RowNo").Cells.Item(pVal.Row).Specific.string

                        subInsertDataIntoInvTransfer(DocEntry, RowNo, objFormInvTransfer, pVal)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing
                    End If
            End Select
        End If

        'UDF Inventory Transfer
        If pVal.FormTypeEx = InvTransferUDF_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormInvTransferUDF = objApplication.Forms.Item(pVal.FormUID)
            End If
        End If

        ' ITEM MASTER
        If pVal.FormTypeEx = ItemMaster_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormItemMaster = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction Then
                        subItemMasterAddObject(objFormItemMaster, True)
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    '    Select Case pVal.ItemUID
                    '        Case "39"
                    '            subItemMasterAddObject(objFormItemMaster, True)
                    '    End Select

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "btnGenItm"
                            If pVal.BeforeAction Then
                                If objFormItemMaster.Items.Item("1").Specific.caption = "Add" Then
                                    subGeneratedItemGroup()
                                Else
                                    objApplication.StatusBar.SetText("Generate Code Only Status Add ~OITM.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo Setnothing
                                End If
                            End If

                    End Select

            End Select
        End If


        'UDF Item Master
        If pVal.FormTypeEx = ItemMasterUDF_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormItemMasterUDF = objApplication.Forms.Item(pVal.FormUID)
            End If
        End If

        'GOOD ISSUE
        If pVal.FormTypeEx = GoodIssue_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormGoodIssue = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction Then

                        'objFormGoodIssue.Items.Item("21").Specific.string = "GW11500"
                        ''objFormGoodIssue.Items.Item("11").Specific.string = "GW11500"
                        subGoodIssueAddObject(objFormGoodIssue, True)

                    End If

                    ' karno lost focus item code hanya transaksi 1(Feed Issue) untuk isi category
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ColUID
                        Case "9", "U_MISFISHQ"
                            If objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
                                objApplication.StatusBar.SetText("you Must Fill Transaction Type", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                GoTo Setnothing
                            Else
                                If objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string = 1 Then
                                    Dim objMatrix As SAPbouiCOM.Matrix
                                    Dim objColumns As SAPbouiCOM.Columns

                                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                                    objColumns = objMatrix.Columns

                                    objColumns.Item("U_MISINFO").Cells.Item(1).Specific.select("2", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                End If
                            End If
                    End Select

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "btnCalc"
                            If pVal.Before_Action = True Then
                                Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim objColumns As SAPbouiCOM.Columns = Nothing
                                Dim Quantity As Double
                                Dim TotalQuantity As Double
                                Dim FishQty As Integer
                                Dim TotalFishQty As Integer

                                objMatrix = objFormGoodIssue.Items.Item("13").Specific
                                objColumns = objMatrix.Columns

                                For i As Integer = 1 To objMatrix.RowCount
                                    If objColumns.Item("9").Cells.Item(i).Specific.value = "" Then
                                        Quantity = 0
                                        TotalQuantity = TotalQuantity + Quantity
                                    Else
                                        Quantity = objColumns.Item("9").Cells.Item(i).Specific.value
                                        TotalQuantity = TotalQuantity + Quantity
                                        If objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.value = "" Then
                                            FishQty = 0
                                            TotalFishQty = TotalFishQty + FishQty
                                        Else
                                            FishQty = objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.value
                                            TotalFishQty = TotalFishQty + FishQty
                                        End If
                                    End If
                                Next

                                objApplication.MessageBox("Total Quantity In Kg = " & TotalQuantity & " AND Total Fish Qty = " & TotalFishQty & " ", 0, "OK")
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                            End If

                        Case "btnCopyGR"
                            If pVal.BeforeAction Then
                                If objFormGoodIssue.Items.Item("1").Specific.caption = "OK" Then
                                    objApplication.StatusBar.SetText("Transaction Already Exists ~OIGE.1.001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo Setnothing
                                Else
                                    If fctFormExist(ListGR_FormId, intFormCountListGR) Then
                                        objApplication.Forms.Item(intFormCountListGR).Select()
                                    Else
                                        subScrPaintListGR()
                                    End If
                                End If
                            End If

                        Case "btnGen"

                            If pVal.BeforeAction = False Then
                                If pVal.ItemUID = "btnGen" Then
                                    'If objFormGoodIssue.Items.Item("btnGen").Specific.caption = "Upload" Then
                                    '    Try
                                    '        'FileName = Me.FindFile()


                                    '    Catch ex As Exception
                                    '        objApplication.MessageBox("Wrong Import ~OIGE.2.001~", 1, "OK")
                                    '    End Try

                                    'Else
                                    If objFormGoodIssue.Items.Item("btnGen").Specific.caption = "Upload" Then
                                        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
                                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                        Dim objColumns As SAPbouiCOM.Columns = Nothing
                                        Dim StrSql As String
                                        Dim Sheet As String
                                        Dim i As Integer
                                        Dim RowTanggal As Integer

                                        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        StrSql = "SELECT U_MISUSRSG UserSign, U_MISROW Row, U_MISCOL Colum, U_MISROWP RowProject FROM [@MIS_EXCELL]"
                                        objRecSet.DoQuery(StrSql)
                                        If objRecSet.RecordCount = 0 Then
                                            Row = 17
                                            Col = 4
                                            RowTanggal = 14
                                            Sheet = "Sheet1"
                                        ElseIf objRecSet.RecordCount = 1 Then
                                            StrSql = "SELECT U_MISSHEET Sheet, U_MISROW Row, U_MISCOL Colum, U_MISROWP RowProject FROM [@MIS_EXCELL]"
                                            objRecSet.DoQuery(StrSql)
                                            Sheet = objRecSet.Fields.Item("Sheet").Value
                                            Row = objRecSet.Fields.Item("Row").Value
                                            Col = objRecSet.Fields.Item("Colum").Value
                                            RowTanggal = objRecSet.Fields.Item("RowProject").Value

                                        Else
                                            StrSql = "SELECT U_MISSHEET Sheet, U_MISROW Row, U_MISCOL Colum, U_MISROWP RowProject FROM [@MIS_EXCELL] WHERE U_MISUSRSG = '" & oCompany.UserSignature & "'"
                                            objRecSet.DoQuery(StrSql)
                                            Sheet = objRecSet.Fields.Item("Sheet").Value
                                            Row = objRecSet.Fields.Item("Row").Value
                                            Col = objRecSet.Fields.Item("Colum").Value
                                            RowTanggal = objRecSet.Fields.Item("RowProject").Value

                                        End If


                                        'Dim oExcel As Excel.Application = Nothing
                                        'Dim oBook As Excel.Workbook = Nothing
                                        'Dim oSheet As Excel.Worksheet = Nothing

                                        Dim oExcel As Microsoft.Office.Interop.Excel.Application = Nothing
                                        Dim oBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
                                        Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                                        FileName = objFormGoodIssueUDF.Items.Item("U_FileName").Specific.value
                                        oExcel = CreateObject("Excel.Application")
                                        oBook = oExcel.Workbooks.Open(FileName)
                                        oSheet = oBook.Worksheets(Sheet)

                                        objMatrix = objFormGoodIssue.Items.Item("13").Specific
                                        objColumns = objMatrix.Columns

                                        If oSheet.Cells(RowTanggal, Col).value = Nothing Then
                                            objApplication.StatusBar.SetText("Date Must Fill", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            GoTo Setnothing

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
                                            oSheet = Nothing
                                            oBook.Close()

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                            oBook = Nothing
                                            oExcel.Quit()

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                            oExcel = Nothing

                                        Else
                                            objFormGoodIssue.Items.Item("9").Specific.string = ClsGlobal.fctFormatDate(oSheet.Cells(RowTanggal, Col).value, oCompany, 1)
                                            objFormGoodIssue.Items.Item("38").Specific.string = ClsGlobal.fctFormatDate(oSheet.Cells(RowTanggal, Col).value, oCompany, 1)
                                        End If
                                        i = 1

                                        objFormGoodIssue.Freeze(True)



                                        Try
                                            Do While oSheet.Cells(Row, Col).value <> ""
                                                objColumns.Item("1").Cells.Item(i).Specific.string = oSheet.Cells(Row, Col).value
                                                objColumns.Item("9").Cells.Item(i).Specific.string = oSheet.Cells(Row, Col + 3).value
                                                objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.string = oSheet.Cells(Row, Col + 4).value
                                                objColumns.Item("U_MISNETID").Cells.Item(i).Specific.string = oSheet.Cells(Row, Col - 1).value
                                                'objColumns.Item("U_MISPROID").Cells.Item(i).Specific.string = oSheet.Cells(RowProject, Col).value
                                                i = i + 1
                                                Row = Row + 1
                                            Loop
                                            'objFormGoodIssue.Items.Item("txtGen").Specific.string = ""
                                            objFormGoodIssue.Items.Item("btnGen").Specific.caption = "Upload"

                                        Catch ex As Exception
                                            objApplication.MessageBox("Check Format Excell ~OIGE.2.002~", 1, "OK")

                                        End Try

                                        objFormGoodIssue.Freeze(False)

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                                        objColumns = Nothing
                                        objMatrix = Nothing
                                        objRecSet = Nothing

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
                                        oSheet = Nothing
                                        oBook.Close()

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
                                        oBook = Nothing
                                        oExcel.Quit()

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
                                        oExcel = Nothing

                                        GC.Collect()

                                    End If
                                End If
                            End If

                            'Case 1
                            '    If pVal.BeforeAction Then
                            '        subValidateGoodIssue()
                            '    End If

                    End Select
            End Select
        End If

        'UDF GOOD ISSUE
        If pVal.FormTypeEx = GoodIssueUDF_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormGoodIssueUDF = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ItemUID
                        Case "U_MISTRXTP"
                            If Not pVal.BeforeAction Then
                                Dim TransactionType As Integer
                                If objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
                                    TransactionType = 0
                                Else
                                    TransactionType = objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string
                                    subLostFocus(TransactionType)
                                End If



                            End If
                    End Select
            End Select

        End If

        ' List GoodReceipt
        If pVal.FormTypeEx = ListGR_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormListGR = objApplication.Forms.Item(pVal.FormUID)
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                blnModalListGr = False
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    'If fctFormExist(GoodIssue_FormId, intRowGoodIssueDetail) Then
                    '    objApplication.Forms.Item(intRowGoodIssueDetail).Select()
                    'Else
                    Dim DocEntry As Integer
                    If (pVal.ColUID = "DocEntry" Or pVal.ColUID = "Number" Or pVal.ColUID = "series" Or pVal.ColUID = "DocDate" Or pVal.ColUID = "RitNo") And (pVal.BeforeAction) And (pVal.Row > 0) Then
                        blnModalListGr = False
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing

                        objMatrix = objFormListGR.Items.Item("MtxGRPO").Specific
                        objColumns = objMatrix.Columns
                        DocEntry = objColumns.Item("DocEntry").Cells.Item(pVal.Row).Specific.string

                        subInsertDataIntoGoodIssue(DocEntry, objFormGoodIssue, pVal)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing
                    End If
            End Select
        End If

        'ListBatch
        If pVal.FormTypeEx = LookUpDistNumber_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormLookUpDistNumber = objApplication.Forms.Item(pVal.FormUID)
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                blnModalLookUpDistNumber = False
            End If
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If (pVal.ColUID = "Batch" Or pVal.ColUID = "NetId" Or pVal.ColUID = "DocDate" Or pVal.ColUID = "RitNo" Or pVal.ColUID = "BoxNo") And (pVal.BeforeAction) And (pVal.Row > 0) Then
                        blnModalLookUpDistNumber = False
                        Dim Batch As String
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing

                        objMatrix = oFormLookUpDistNumber.Items.Item("MtxBatch").Specific
                        objColumns = objMatrix.Columns
                        Batch = objColumns.Item("Batch").Cells.Item(pVal.Row).Specific.string

                        oFormLookUpDistNumber.Close()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing

                        Dim objMatrixBatch As SAPbouiCOM.Matrix = Nothing
                        Dim objColumnsBatch As SAPbouiCOM.Columns = Nothing

                        'objMatrixBatch = objApplication.Forms.Item(LookUpBatch_FormId).Items.Item("3").Specific

                        objMatrixBatch = objFormBatch.Items.Item("3").Specific
                        objColumnsBatch = objMatrixBatch.Columns
                        objColumnsBatch.Item("2").Cells.Item(1).Specific.string = Batch
                        'End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixBatch)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsBatch)
                        objMatrixBatch = Nothing
                        objColumnsBatch = Nothing

                    End If
            End Select
        End If

        ' BatchForm
        If pVal.FormTypeEx = Batch_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormBatch = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                    Dim objColumns As SAPbouiCOM.Columns = Nothing
                    Dim objMatrixBatch As SAPbouiCOM.Matrix = Nothing
                    Dim objColumnsBatch As SAPbouiCOM.Columns = Nothing
                    Dim objMatrixListBatch As SAPbouiCOM.Matrix = Nothing
                    Dim objColumnsListBatch As SAPbouiCOM.Columns = Nothing
                    Dim ItemCode As String
                    Dim ProjectId As String
                    Dim Quantity As Double
                    Dim DelDate As Date
                    Dim Tahun As Integer
                    Dim Bulan As Integer
                    Dim StrBulan As String
                    Dim Tanggal As Integer
                    Dim StrTanggal As String
                    Dim StrFullDate As String
                    Dim RitNo As String
                    Dim BoxNo As String
                    Dim Batch As String
                    Dim Shift As String
                    'Dim Country As String
                    'Dim Region As String
                    'karno

                    If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 1 Then
                        If objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.value = "" Then
                            objFormBatch.Close()
                            objApplication.MessageBox("You must fill out the Rit No first ", 1, "OK")
                            GoTo Setnothing
                        Else
                            RitNo = objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.selected.value

                            objMatrixListBatch = objFormBatch.Items.Item("35").Specific
                            objColumnsListBatch = objMatrixListBatch.Columns

                            For Line As Integer = 1 To objMatrixListBatch.RowCount
                                If objMatrixListBatch.IsRowSelected(Line) = True Then
                                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                    objColumns = objMatrix.Columns
                                    ItemCode = objColumns.Item("1").Cells.Item(Line).Specific.string
                                    If objColumns.Item("U_MISPROID").Cells.Item(Line).Specific.string = "" Then
                                        objFormBatch.Close()
                                        objApplication.MessageBox("You must fill out the Project Id first ", 1, "OK")
                                        objColumns.Item("U_MISPROID").Cells.Item(Line).Click()
                                        GoTo Setnothing
                                    Else
                                        ProjectId = objColumns.Item("U_MISPROID").Cells.Item(Line).Specific.string
                                        'If objColumns.Item("U_MISBOXNO").Cells.Item(Line).Specific.value = "" Then
                                        '    objFormBatch.Close()
                                        '    objApplication.MessageBox("You must fill out the Box No first ", 1, "OK")
                                        '    objColumns.Item("U_MISBOXNO").Cells.Item(Line).Click()
                                        '    GoTo Setnothing
                                        'Else
                                        If objColumns.Item("U_MISBATCH").Cells.Item(Line).Specific.string = "" Then
                                            objFormBatch.Close()
                                            objApplication.StatusBar.SetText("You must fill out the Batch first ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objColumns.Item("U_MISBATCH").Cells.Item(Line).Click()
                                            GoTo Setnothing
                                        Else
                                            'BoxNo = objColumns.Item("U_MISBOXNO").Cells.Item(Line).Specific.selected.value
                                            Quantity = objColumns.Item("9").Cells.Item(Line).Specific.string
                                            objMatrixBatch = objFormBatch.Items.Item("3").Specific
                                            objColumnsBatch = objMatrixBatch.Columns
                                            objColumnsBatch.Item("2").Cells.Item(1).Specific.string = ProjectId + "." + RitNo + "." + BoxNo
                                            'objColumnsBatch.Item("5").Cells.Item(1).Specific.string = Quantity
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixBatch)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsBatch)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                            objMatrixBatch = Nothing
                                            objColumnsBatch = Nothing
                                            objMatrix = Nothing
                                            objColumns = Nothing

                                            Line = Line + 1
                                        End If
                                        'End If
                                    End If
                                End If
                            Next
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixListBatch)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsListBatch)
                            objMatrixListBatch = Nothing
                            objColumnsListBatch = Nothing


                        End If
                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 2 Or objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 3 Then
                        objMatrixListBatch = objFormBatch.Items.Item("35").Specific
                        objColumnsListBatch = objMatrixListBatch.Columns

                        For Line As Integer = 1 To objMatrixListBatch.RowCount
                            If objMatrixListBatch.IsRowSelected(Line) = True Then
                                objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                objColumns = objMatrix.Columns

                                If objColumns.Item("U_MISBATCH").Cells.Item(Line).Specific.string = "" Then
                                    objApplication.StatusBar.SetText("U Must Fill Batch Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                    objFormBatch.Close()
                                    GoTo Setnothing
                                Else
                                    Batch = objColumns.Item("U_MISBATCH").Cells.Item(Line).Specific.string

                                    objMatrixBatch = objFormBatch.Items.Item("3").Specific
                                    objColumnsBatch = objMatrixBatch.Columns

                                    objColumnsBatch.Item("2").Cells.Item(1).Specific.string = Batch

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixBatch)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsBatch)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                    objMatrixBatch = Nothing
                                    objColumnsBatch = Nothing
                                    objMatrix = Nothing
                                    objColumns = Nothing
                                End If
                            End If
                        Next

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixListBatch)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsListBatch)
                        objMatrixListBatch = Nothing
                        objColumnsListBatch = Nothing

                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 4 Then

                        objMatrixListBatch = objFormBatch.Items.Item("35").Specific
                        objColumnsListBatch = objMatrixListBatch.Columns
                        'Dim Shift As String
                        Dim Project As String

                        If objFormGoodReceipt.Items.Item("11").Specific.string = "" Then 'comments
                            objApplication.StatusBar.SetText("U Must Using Button Copy From Good Issue To Get Project Code ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objFormBatch.Close()
                            GoTo Setnothing
                        Else

                            If objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Specific.string = "" Then
                                objApplication.StatusBar.SetText("You Must Fill Shift", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objFormBatch.Close()
                                objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Click()
                                GoTo Setnothing
                            Else


                                Shift = objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Specific.string
                                Project = objFormGoodReceipt.Items.Item("11").Specific.string

                                DelDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormGoodReceipt.Items.Item("9").Specific.string, 5))
                                Tahun = Right(DelDate.Year, 2)
                                Bulan = DelDate.Month
                                If Len(Trim(Bulan)) = 1 Then
                                    StrBulan = "0" + CStr(Bulan)
                                Else
                                    StrBulan = CStr(Bulan)
                                End If
                                Tanggal = DelDate.Day
                                If Len(Trim(Tanggal)) = 1 Then
                                    StrTanggal = "0" + CStr(Tanggal)
                                Else
                                    StrTanggal = CStr(Tanggal)
                                End If
                                StrFullDate = CStr(Tahun) + StrBulan + StrTanggal

                                Batch = Project + "." + StrFullDate + "." + Shift

                                For Line As Integer = 1 To objMatrixListBatch.RowCount
                                    If objMatrixListBatch.IsRowSelected(Line) = True Then
                                        objMatrixBatch = objFormBatch.Items.Item("3").Specific
                                        objColumnsBatch = objMatrixBatch.Columns

                                        objColumnsBatch.Item("2").Cells.Item(1).Specific.string = Batch

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixBatch)
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsBatch)
                                        objMatrixBatch = Nothing
                                        objColumnsBatch = Nothing
                                    End If
                                Next

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixListBatch)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsListBatch)
                                objMatrixListBatch = Nothing
                                objColumnsListBatch = Nothing
                            End If
                        End If
                    Else
                        objApplication.MessageBox("Generate Data only For Harvesting, U Can Input Manual", 0, "OK")
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction Then
                        subBatchAddObject(objFormBatch, True)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "btnBatch"

                            If pVal.BeforeAction Then
                                Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim objColumns As SAPbouiCOM.Columns = Nothing
                                Dim ItemCode As String
                                Dim StrSql As String
                                Dim ItemGroup As String
                                Dim objRecSet As SAPbobsCOM.Recordset = Nothing
                                objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                objMatrix = objFormBatch.Items.Item("35").Specific
                                objColumns = objMatrix.Columns

                                ItemCode = objColumns.Item("5").Cells.Item(1).Specific.string()


                                'StrSql = "SELECT T0.[ItmsGrpCod] ItemGroup FROM OITB T0 INNER JOIN OITM T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod Where T1.ItemCode = '" & ItemCode & "'"

                                StrSql = "SELECT T0.[ItmsGrpNam] ItemGroup FROM OITB T0 INNER JOIN OITM T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod Where T1.ItemCode = '" & ItemCode & "'"
                                objRecSet.DoQuery(StrSql)

                                ItemGroup = objRecSet.Fields.Item("ItemGroup").Value

                                'If ItemGroup >= 101 And ItemGroup <= 104 Then
                                If ItemGroup = "RM Feedmill" Or ItemGroup = "RM Feed" _
                                    Or ItemGroup = "RM Biomass" Or ItemGroup = "RM Fingerling" Then
                                    subScrPaintBatchRM()
                                ElseIf ItemGroup = "Finish Goods" Then
                                    SubScrPaintBatchFG()
                                    'objApplication.MessageBox("Finish Good", 1, "OK")
                                Else
                                    objApplication.MessageBox("No Data ~BTCH.1.0001~", 1, "OK")
                                End If

                                'subFormLoadDistNumber(pVal.Row)

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                                objMatrix = Nothing
                                objColumns = Nothing
                            End If
                    End Select

                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    '    If pVal.BeforeAction = False Then
                    '        Select Case pVal.ColUID
                    '            Case 2
                    '                Dim DocEntry As Integer
                    '                DocEntry = objFormGoodReceipt.Items.Item("txtBatch").Specific.string
                    '                subFormLoadDistNumber(DocEntry, pVal.Row)
                    '        End Select
                    '    End If
            End Select
        End If

        'BatchRM_FormId
        If pVal.FormTypeEx = BatchRM_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormBatchRM = objApplication.Forms.Item(pVal.FormUID)
            End If


            Select Case pVal.ItemUID
                Case "btnGen"
                    If pVal.BeforeAction Then
                        If Not fctValidateBatchRM(oFormBatchRM, BubbleEvent) Then GoTo Setnothing
                        subGeneratedBatchRM()
                    End If

            End Select
        End If

        'BatchFG_FormId
        If pVal.FormTypeEx = BatchFG_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormBatchFG = objApplication.Forms.Item(pVal.FormUID)
            End If


            Select Case pVal.ItemUID
                Case "btnGen"
                    If pVal.BeforeAction Then
                        If Not fctValidateBatchFG(oFormBatchFG, BubbleEvent) Then GoTo Setnothing
                        subGeneratedBatchFG()
                    End If

            End Select
        End If

        ' GOOD RECEIPT 
        If pVal.FormTypeEx = GoodReceipt_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormGoodReceipt = objApplication.Forms.Item(pVal.FormUID)
            End If


            'GOOD RECEIPT FORM LOAD
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.BeforeAction Then
                        subGoodReceiptAddObject(objFormGoodReceipt, True)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ColUID
                        Case "1"
                            If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                                objApplication.StatusBar.SetText("You Must Fill Transaction Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                GoTo Setnothing
                            Else
                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
                                    objApplication.StatusBar.SetText("You Must Fill Transaction Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo Setnothing
                                Else
                                    If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 3 Then
                                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                        Dim objColumns As SAPbouiCOM.Columns = Nothing

                                        objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                        objColumns = objMatrix.Columns
                                        If objColumns.Item("1").Cells.Item(pVal.Row).Specific.string = "" Then
                                            objApplication.StatusBar.SetText("You Must Input Item code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            GoTo Setnothing
                                        Else
                                            objColumns.Item("15").Cells.Item(pVal.Row).Specific.string = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string
                                            If objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.value = "" Then
                                                objApplication.StatusBar.SetText("You Must Input Rit No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                GoTo Setnothing
                                            Else
                                                Dim project As String
                                                Dim ritno As String
                                                project = objFormGoodReceipt.Items.Item("11").Specific.string
                                                ritno = objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.selected.value
                                                objColumns.Item("U_MISBATCH").Cells.Item(pVal.Row).Specific.string = project + "." + ritno
                                            End If
                                        End If

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                        objMatrix = Nothing
                                        objColumns = Nothing
                                    End If
                                End If
                            End If



                        Case "9", "U_MISINFO"
                            If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
                                objApplication.StatusBar.SetText("You Must Fill Transaction Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                GoTo Setnothing
                            Else
                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 3 Or objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 4 Then
                                    Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim objColumns As SAPbouiCOM.Columns = Nothing

                                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                    objColumns = objMatrix.Columns
                                    If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = "" Then
                                        objApplication.StatusBar.SetText("You Must Fill Transaction Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        GoTo Setnothing
                                    Else
                                        objColumns.Item("15").Cells.Item(pVal.Row).Specific.string = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string
                                        If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 3 Then
                                            If objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.value = "" Then
                                                objApplication.StatusBar.SetText("You Must Input Rit No", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                GoTo Setnothing
                                            Else
                                                Dim project As String
                                                Dim ritno As String
                                                project = objFormGoodReceipt.Items.Item("11").Specific.string
                                                ritno = objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.selected.value
                                                objColumns.Item("U_MISPROID").Cells.Item(pVal.Row).Specific.string = project
                                                objColumns.Item("U_MISBATCH").Cells.Item(pVal.Row).Specific.string = project + "." + ritno

                                                If objColumns.Item("U_MISINFO").Cells.Item(pVal.Row).Specific.value = "" Then
                                                    objApplication.StatusBar.SetText("You Must Fill Category", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                                    objMatrix = Nothing
                                                    objColumns = Nothing
                                                    GoTo Setnothing
                                                ElseIf objColumns.Item("U_MISINFO").Cells.Item(pVal.Row).Specific.value <> 1 Then
                                                    'ElseIf objColumns.Item("U_MISINFO").Cells.Item(pVal.Row).Specific.selected.value <> 1 Then
                                                    objColumns.Item("10").Cells.Item(pVal.Row).Specific.value = 1
                                                    'Else
                                                    '    objColumns.Item("10").Cells.Item(pVal.Row).Specific.value = objColumns.Item("U_MISGISTV").Cells.Item(pVal.Row).Specific.value / CInt(objColumns.Item("9").Cells.Item(pVal.Row).Specific.value)
                                                    '    objColumns.Item("14").Cells.Item(pVal.Row).Specific.value = objColumns.Item("U_MISGISTV").Cells.Item(pVal.Row).Specific.value

                                                End If
                                            End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                            objMatrix = Nothing
                                            objColumns = Nothing
                                        End If
                                    End If
                                End If
                            End If
                    End Select

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID

                        Case "btnToGI"
                            If pVal.BeforeAction Then
                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = "1" Then
                                    Dim oGoodIssue As SAPbobsCOM.Documents
                                    Dim oUDFGoodIssue As SAPbobsCOM.UserFields
                                    Dim oRecset As SAPbobsCOM.Recordset
                                    Dim oRecset1 As SAPbobsCOM.Recordset
                                    Dim oRecsetSeries As SAPbobsCOM.Recordset
                                    Dim oRecsetAccount As SAPbobsCOM.Recordset
                                    Dim strSqlAccount As String
                                    Dim strSqlSeries As String
                                    Dim strSql1 As String
                                    Dim strSql As String
                                    Dim objMatrix As SAPbouiCOM.Matrix
                                    Dim objcolumns As SAPbouiCOM.Columns
                                    Dim lngResult As Long
                                    Dim IntType As Integer
                                    Dim Tipe As Integer

                                    oGoodIssue = oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit)
                                    oUDFGoodIssue = oGoodIssue.UserFields
                                    oRecset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    oRecset1 = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    oRecsetSeries = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    oRecsetAccount = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                                    Tipe = objApplication.MessageBox("Transaction Type", 1, "Generate GI", "Revisi", "Cancel")


                                    If objFormGoodReceiptUDF.Items.Item("U_MISTargetNum").Specific.value <> "" Then
                                        objApplication.MessageBox("Transaction Sudah Digerate No Dokumen" + objFormGoodReceiptUDF.Items.Item("U_MISTargetNum").Specific.value)
                                    Else
                                        'Generate Good Issue
                                        If Tipe = 1 Then
                                            IntType = objApplication.MessageBox("Transaction Type Fish Issue To", 1, "Processing", "Biomass", "Project")

                                            strSqlSeries = "SELECT Series, NextNumber FROM NNM1 WHERE InitialNum = " & Left(objFormGoodReceipt.Items.Item("7").Specific.value, 4) + "00001" & " AND ObjectCode = 60"
                                            oRecsetSeries.DoQuery(strSqlSeries)

                                            If oRecsetSeries.RecordCount = 1 Then
                                                oGoodIssue.Series = oRecsetSeries.Fields.Item("Series").Value
                                            Else
                                                objApplication.MessageBox("Wrong Numbering")
                                            End If

                                            oGoodIssue.DocDate = ClsGlobal.fctFormatDate(objFormGoodReceipt.Items.Item("9").Specific.string, oCompany, 1)
                                            oGoodIssue.Reference2 = objFormGoodReceipt.Items.Item("21").Specific.string

                                            'Processing
                                            If IntType = 1 Then
                                                oUDFGoodIssue.Fields.Item("U_MISTRXTP").Value = 5
                                                oUDFGoodIssue.Fields.Item("U_MISTRXNM").Value = "Fish Transfer To PP"
                                                'biomass
                                            ElseIf IntType = 2 Then
                                                oUDFGoodIssue.Fields.Item("U_MISTRXTP").Value = 3
                                                oUDFGoodIssue.Fields.Item("U_MISTRXNM").Value = "Fish Transfer To Biomass"
                                                'project
                                            ElseIf IntType = 3 Then
                                                oUDFGoodIssue.Fields.Item("U_MISTRXTP").Value = 2
                                                oUDFGoodIssue.Fields.Item("U_MISTRXNM").Value = "Fish Issue To Project"
                                            End If

                                            If IntType = 1 Or IntType = 2 Then
                                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value = "" Then
                                                    objApplication.MessageBox("Transaction Warehouse Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISDESTW").Specific.value = "" Then
                                                    objApplication.MessageBox("Destination Warehouse Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Specific.value = "" Then
                                                    objApplication.MessageBox("Driver name Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Specific.value = "" Then
                                                    objApplication.MessageBox("Assistant Driver name Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Specific.value = "" Then
                                                    objApplication.MessageBox("Licence No Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Specific.value = "" Then
                                                    objApplication.MessageBox("Supervisor Id Belum Terisi")
                                                ElseIf objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.value = "" Then
                                                    objApplication.MessageBox("Delivery Time Belum Terisi")
                                                    'ElseIf objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.Value = "" Then
                                                    '    objApplication.MessageBox("Rit No Belum Terisi")
                                                Else
                                                    oUDFGoodIssue.Fields.Item("U_MISTRXWH").Value = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISDESTW").Value = objFormGoodReceiptUDF.Items.Item("U_MISDESTW").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISDRVNM").Value = objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISASDRV").Value = objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISLICNO").Value = objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISSPVID").Value = objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISDELTM").Value = CDate("2011-01-01 " + Left(objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.value, 2) + ":" + Right(objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.value, 2) + ":00")
                                                    oUDFGoodIssue.Fields.Item("U_MISRITNO").Value = objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISREFFD").Value = objFormGoodReceipt.Items.Item("7").Specific.value

                                                    strSql = "SELECT T2.BatchNum, T1.Quantity, T1.ItemCode, T1.WhsCode " & _
                                                         "FROM OIGN T0 " & _
                                                         "INNER JOIN IGN1 T1 " & _
                                                         "ON T0.DocEntry = T1.DocEntry " & _
                                                         "INNER JOIN IBT1 T2 " & _
                                                         "ON T1.ItemCode = T2.ItemCode " & _
                                                         "AND T1.WhsCode = T2.WhsCode " & _
                                                         "AND T0.ObjType = T2.BaseType " & _
                                                         "AND T0.DocEntry = T2.BaseEntry " & _
                                                         "AND T1.LineNum = T2.BaseLinNum " & _
                                                         "WHERE T0.DocNum = '" & objFormGoodReceipt.Items.Item("7").Specific.Value & "' "

                                                    oRecset.DoQuery(strSql)

                                                    If oRecset.RecordCount <= 0 Then
                                                        objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                    Else
                                                        objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                                        objcolumns = objMatrix.Columns

                                                        strSql1 = "SELECT T0.Quantity From OIBT T0 WHERE T0.BatchNum = '" & oRecset.Fields.Item("BatchNum").Value & "'"

                                                        oRecset1.DoQuery(strSql1)
                                                        If oRecset1.RecordCount = 0 Then
                                                            objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                        Else
                                                            If oRecset1.Fields.Item("Quantity").Value > 0 Then
                                                                For i As Integer = 0 To oRecset.RecordCount - 1
                                                                    oGoodIssue.Lines.SetCurrentLine(i)
                                                                    oGoodIssue.Lines.ItemCode = objcolumns.Item("1").Cells.Item(i + 1).Specific.Value
                                                                    oGoodIssue.Lines.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISFISHQ").Value = objcolumns.Item("U_MISFISHQ").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.WarehouseCode = objcolumns.Item("15").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISNETID").Value = objcolumns.Item("U_MISNETID").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISPROID").Value = objcolumns.Item("U_MISPROID").Cells.Item(i + 1).Specific.value

                                                                    strSqlAccount = "SELECT T1.AcctCode Account FROM OITW T0 INNER JOIN OACT T1 ON REPLACE(T0.[U_GIT],'-','') = REPLACE(T1.FormatCode,'-','')" & _
                                                                             "WHERE T0.WhsCode = '" & objFormGoodReceiptUDF.Items.Item("U_MISDESTW").Specific.String & "' " & _
                                                                             "AND ItemCode = '" & objcolumns.Item("1").Cells.Item(i + 1).Specific.Value & "'"

                                                                    oRecsetAccount.DoQuery(strSqlAccount)

                                                                    If oRecsetAccount.RecordCount = 0 Then
                                                                        objApplication.MessageBox("Account Not Found, Please check GL Account")
                                                                    Else
                                                                        oGoodIssue.Lines.AccountCode = oRecsetAccount.Fields.Item("Account").Value
                                                                    End If

                                                                    oGoodIssue.Lines.BatchNumbers.SetCurrentLine(i)
                                                                    oGoodIssue.Lines.BatchNumbers.BatchNumber = oRecset.Fields.Item("BatchNum").Value
                                                                    oGoodIssue.Lines.BatchNumbers.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value

                                                                    oGoodIssue.Lines.BatchNumbers.Add()
                                                                    oGoodIssue.Lines.Add()

                                                                    oRecset.MoveNext()
                                                                Next
                                                                oGoodIssue.Add()

                                                            End If
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objcolumns)
                                                            objMatrix = Nothing
                                                            objcolumns = Nothing
                                                        End If
                                                    End If
                                                End If
                                                'Project
                                            ElseIf IntType = 3 Then
                                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value = "" Then
                                                    objApplication.MessageBox("Transaction Warehouse Belum Terisi")
                                                Else
                                                    oUDFGoodIssue.Fields.Item("U_MISTRXWH").Value = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value
                                                    oUDFGoodIssue.Fields.Item("U_MISREFFD").Value = objFormGoodReceipt.Items.Item("7").Specific.value

                                                    strSql = "SELECT T2.BatchNum, T1.Quantity, T1.ItemCode, T1.WhsCode " & _
                                                         "FROM OIGN T0 " & _
                                                         "INNER JOIN IGN1 T1 " & _
                                                         "ON T0.DocEntry = T1.DocEntry " & _
                                                         "INNER JOIN IBT1 T2 " & _
                                                         "ON T1.ItemCode = T2.ItemCode " & _
                                                         "AND T1.WhsCode = T2.WhsCode " & _
                                                         "AND T0.ObjType = T2.BaseType " & _
                                                         "AND T0.DocEntry = T2.BaseEntry " & _
                                                         "AND T1.LineNum = T2.BaseLinNum " & _
                                                         "WHERE T0.DocNum = '" & objFormGoodReceipt.Items.Item("7").Specific.Value & "' "

                                                    oRecset.DoQuery(strSql)

                                                    If oRecset.RecordCount >= 0 Then
                                                        objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                                        objcolumns = objMatrix.Columns

                                                        strSql1 = "SELECT T0.Quantity From OIBT T0 WHERE T0.BatchNum = '" & oRecset.Fields.Item("BatchNum").Value & "'"

                                                        oRecset1.DoQuery(strSql1)
                                                        If oRecset1.RecordCount > 0 Then
                                                            If oRecset1.Fields.Item("Quantity").Value > 0 Then
                                                                For i As Integer = 0 To oRecset.RecordCount - 1
                                                                    oGoodIssue.Lines.SetCurrentLine(i)
                                                                    oGoodIssue.Lines.ItemCode = objcolumns.Item("1").Cells.Item(i + 1).Specific.Value
                                                                    oGoodIssue.Lines.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISFISHQ").Value = objcolumns.Item("U_MISFISHQ").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.WarehouseCode = objcolumns.Item("15").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISNETID").Value = objcolumns.Item("U_MISNETID").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISPROID").Value = objcolumns.Item("U_MISPROID").Cells.Item(i + 1).Specific.value
                                                                    oGoodIssue.Lines.UserFields.Fields.Item("U_MISBATCH").Value = objcolumns.Item("U_MISBATCH").Cells.Item(i + 1).Specific.value
                                                                    strSqlAccount = "SELECT T1.AcctCode Account FROM OITW T0 INNER JOIN OACT T1 ON REPLACE(T0.[U_WIP],'-','') = REPLACE(T1.FormatCode,'-','') " & _
                                                                             "WHERE T0.WhsCode = '" & objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.String & "' " & _
                                                                             "AND ItemCode = '" & objcolumns.Item("1").Cells.Item(i + 1).Specific.Value & "'"

                                                                    oRecsetAccount.DoQuery(strSqlAccount)

                                                                    If oRecsetAccount.RecordCount > 0 Then
                                                                        oGoodIssue.Lines.AccountCode = oRecsetAccount.Fields.Item("Account").Value
                                                                    Else
                                                                        objApplication.MessageBox("Account Not Found, Please check GL Account")
                                                                    End If

                                                                    oGoodIssue.Lines.BatchNumbers.SetCurrentLine(i)
                                                                    oGoodIssue.Lines.BatchNumbers.BatchNumber = oRecset.Fields.Item("BatchNum").Value
                                                                    oGoodIssue.Lines.BatchNumbers.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value

                                                                    oGoodIssue.Lines.BatchNumbers.Add()
                                                                    oGoodIssue.Lines.Add()

                                                                    oRecset.MoveNext()
                                                                Next
                                                                oGoodIssue.Add()
                                                                objApplication.MessageBox("No Dokumen Good Issue = '" & oRecsetSeries.Fields.Item("NextNumber").Value & "'", 1)
                                                            Else
                                                                objApplication.MessageBox("Batch Number Already using, U Must Input Good Issue Manual", 1, "OK")
                                                            End If
                                                        Else
                                                            'Batch number tidak ditemukan
                                                            objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                        End If

                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objcolumns)
                                                        objMatrix = Nothing
                                                        objcolumns = Nothing
                                                    Else
                                                        objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                    End If
                                                End If

                                            End If

                                            lngResult = oGoodIssue.Add
                                            If lngResult <> 0 Then
                                                objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                            Else
                                                objApplication.MessageBox("No Dokumen Good Issue = '" & oRecsetSeries.Fields.Item("NextNumber").Value & "'", 1)
                                            End If


                                            'Revisi
                                        ElseIf Tipe = 2 Then

                                            strSqlSeries = "SELECT Series, NextNumber FROM NNM1 WHERE InitialNum = " & Left(objFormGoodReceipt.Items.Item("7").Specific.value, 4) + "00001" & " AND ObjectCode = 60"
                                            oRecsetSeries.DoQuery(strSqlSeries)

                                            If oRecsetSeries.RecordCount = 1 Then
                                                oGoodIssue.Series = oRecsetSeries.Fields.Item("Series").Value
                                            Else
                                                objApplication.MessageBox("Wrong Numbering")
                                            End If


                                            oGoodIssue.DocDate = ClsGlobal.fctFormatDateSave(oCompany, objFormGoodReceipt.Items.Item("9").Specific.value, 4) '+ " 00:00:00"
                                            oGoodIssue.Reference2 = objFormGoodReceipt.Items.Item("21").Specific.string
                                            oUDFGoodIssue.Fields.Item("U_MISTRXTP").Value = 8

                                            If objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value = "" Then
                                                objApplication.MessageBox("Transaction Warehouse Belum Terisi")
                                            Else
                                                oUDFGoodIssue.Fields.Item("U_MISTRXWH").Value = objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.value
                                                oUDFGoodIssue.Fields.Item("U_MISREFFD").Value = objFormGoodReceipt.Items.Item("7").Specific.value

                                                strSql = "SELECT T2.BatchNum, T1.Quantity, T1.ItemCode, T1.WhsCode " & _
                                                     "FROM OIGN T0 " & _
                                                     "INNER JOIN IGN1 T1 " & _
                                                     "ON T0.DocEntry = T1.DocEntry " & _
                                                     "INNER JOIN IBT1 T2 " & _
                                                     "ON T1.ItemCode = T2.ItemCode " & _
                                                     "AND T1.WhsCode = T2.WhsCode " & _
                                                     "AND T0.ObjType = T2.BaseType " & _
                                                     "AND T0.DocEntry = T2.BaseEntry " & _
                                                     "AND T1.LineNum = T2.BaseLinNum " & _
                                                     "WHERE T0.DocNum = '" & objFormGoodReceipt.Items.Item("7").Specific.Value & "' "

                                                oRecset.DoQuery(strSql)

                                                If oRecset.RecordCount >= 0 Then
                                                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                                    objcolumns = objMatrix.Columns

                                                    strSql1 = "SELECT T0.Quantity From OIBT T0 WHERE T0.BatchNum = '" & oRecset.Fields.Item("BatchNum").Value & "'"

                                                    oRecset1.DoQuery(strSql1)
                                                    If oRecset1.RecordCount >= 0 Then
                                                        If oRecset1.Fields.Item("Quantity").Value > 0 Then
                                                            For i As Integer = 0 To oRecset.RecordCount - 1
                                                                oGoodIssue.Lines.SetCurrentLine(i)
                                                                oGoodIssue.Lines.ItemCode = objcolumns.Item("1").Cells.Item(i + 1).Specific.Value
                                                                oGoodIssue.Lines.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value
                                                                oGoodIssue.Lines.UserFields.Fields.Item("U_MISFISHQ").Value = objcolumns.Item("U_MISFISHQ").Cells.Item(i + 1).Specific.value
                                                                oGoodIssue.Lines.WarehouseCode = objcolumns.Item("15").Cells.Item(i + 1).Specific.value
                                                                oGoodIssue.Lines.UserFields.Fields.Item("U_MISNETID").Value = objcolumns.Item("U_MISNETID").Cells.Item(i + 1).Specific.value
                                                                oGoodIssue.Lines.UserFields.Fields.Item("U_MISPROID").Value = objcolumns.Item("U_MISPROID").Cells.Item(i + 1).Specific.value

                                                                strSqlAccount = "SELECT T1.AcctCode Account FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry " & _
                                                                                "WHERE T0.DocNum = '" & objFormGoodReceipt.Items.Item("7").Specific.String & "' "

                                                                oRecsetAccount.DoQuery(strSqlAccount)

                                                                If oRecsetAccount.RecordCount <= 0 Then
                                                                    objApplication.MessageBox("Account Not Found, Please check GL Account")
                                                                Else
                                                                    oGoodIssue.Lines.AccountCode = oRecsetAccount.Fields.Item("Account").Value
                                                                End If

                                                                oGoodIssue.Lines.BatchNumbers.SetCurrentLine(i)
                                                                oGoodIssue.Lines.BatchNumbers.BatchNumber = oRecset.Fields.Item("BatchNum").Value
                                                                oGoodIssue.Lines.BatchNumbers.Quantity = objcolumns.Item("9").Cells.Item(i + 1).Specific.value

                                                                oGoodIssue.Lines.BatchNumbers.Add()
                                                                oGoodIssue.Lines.Add()

                                                                oRecset.MoveNext()
                                                            Next
                                                            oGoodIssue.Add()
                                                            objApplication.MessageBox("generate Good Issue Successfull, no Dokumen = '" & oRecsetSeries.Fields.Item("NextNumber").Value & "'", 1)
                                                        Else
                                                            objApplication.MessageBox("Batch Already Using, U Must Input good issue manual", 1, "OK")
                                                        End If
                                                    Else
                                                        objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                    End If
                                                Else
                                                    objApplication.MessageBox(oCompany.GetLastErrorDescription)
                                                End If
                                            End If
                                        End If
                                        End If

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGoodIssue)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDFGoodIssue)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset1)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecsetSeries)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecsetAccount)

                                    oGoodIssue = Nothing
                                    oRecset = Nothing
                                    oRecset1 = Nothing
                                    oRecsetSeries = Nothing
                                    oRecsetAccount = Nothing

                                    Else
                                        Dim i As Integer
                                        i = objApplication.MessageBox("You Must Input Manual", 1, "OK", "Cancel")
                                        If i = 1 Then
                                            objApplication.ActivateMenuItem("3079")
                                        End If
                                    End If
                            End If

                        Case "btnCalc"
                            If pVal.Before_Action = True Then
                                Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim objColumns As SAPbouiCOM.Columns = Nothing
                                Dim Quantity As Double
                                Dim TotalQuantity As Double
                                Dim FishQty As Integer
                                Dim TotalFishQty As Integer

                                objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                objColumns = objMatrix.Columns

                                For i As Integer = 1 To objMatrix.RowCount
                                    If objColumns.Item("9").Cells.Item(i).Specific.value = "" Then
                                        Quantity = 0
                                        TotalQuantity = TotalQuantity + Quantity
                                    Else
                                        Quantity = objColumns.Item("9").Cells.Item(i).Specific.value
                                        TotalQuantity = TotalQuantity + Quantity
                                        If objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.value = "" Then
                                            FishQty = 0
                                            TotalFishQty = TotalFishQty + FishQty
                                        Else
                                            FishQty = objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.value
                                            TotalFishQty = TotalFishQty + FishQty
                                        End If
                                    End If
                                Next

                                objApplication.MessageBox("Total Quantity In Kg = " & TotalQuantity & " AND Total Fish Qty = " & TotalFishQty & " ", 0, "OK")
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                            End If

                        Case "btnCopyGI"
                            If pVal.BeforeAction Then
                                If objFormGoodReceipt.Items.Item("1").Specific.caption = "OK" Then
                                    objApplication.StatusBar.SetText("Transaction Already Exists ~OIGN.1.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo Setnothing
                                Else
                                    If fctFormExist(ListGI_FormId, intFormCountListGI) Then
                                        objApplication.Forms.Item(intFormCountListGI).Select()
                                    Else
                                        subScrPaintListGI()
                                    End If
                                End If
                            End If

                        Case "1"
                            '    If pVal.BeforeAction Then
                            '        objApplication.MessageBox("HahA", 1, "OK", "Cancel")
                            '    End If

                            'Case "btnCal"


                            If pVal.BeforeAction Then
                                If objFormGoodReceipt.Mode = BoFormMode.fm_ADD_MODE Then
                                    Dim ReferDocument As Integer
                                    Dim ConvertionFish As Double
                                    Dim TotalConvertionFish As Double
                                    Dim TotalFishQty As Integer
                                    Dim QtyInKg As Double
                                    Dim TotalGI As Double
                                    Dim Flag As String
                                    Dim TotalGRNOTFirstGrade As Double
                                    Dim TotalGRFishQty As Integer
                                    Dim EstimasiFish As Integer
                                    Dim MortalityQty As Integer
                                    Dim TotalProjectCost As Double
                                    Dim TotalGoodReceiptCost As Double
                                    Dim StrSql As String
                                    Dim LineTotal As Double
                                    Dim intMsg As Integer
                                    'Dim netId As String
                                    'Dim Tanggal As Date
                                    Dim objRecSet As SAPbobsCOM.Recordset = Nothing
                                    Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                                    Dim objColumns As SAPbouiCOM.Columns = Nothing

                                    If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.value = "" Then
                                        objApplication.StatusBar.SetText("Please Input Transaction Type ~OIGN.2.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        GoTo Setnothing
                                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 1 Then

                                        If objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "" Then
                                            intMsg = objApplication.MessageBox("Are You Sure To Last Harvest? ", 1, "Yes", "No")
                                            If intMsg = 1 Then
                                                objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "Y"
                                            Else
                                                objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "N"
                                            End If
                                            'objApplication.StatusBar.SetText("Please Input Last Harvesting Batch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                            objColumns = objMatrix.Columns

                                            'If objColumns.Item("U_MISNETID").Cells.Item(1).Specific.string = "" Then
                                            '    'objApplication.StatusBar.SetText("Net Id Must Be Fill", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            '    objApplication.MessageBox("Net Id Must Fill ~OIGN.2.0002~", 1, "OK")
                                            '    GoTo Setnothing
                                            'Else

                                            'netId = objColumns.Item("U_MISNETID").Cells.Item(1).Specific.string
                                            'Tanggal = ClsGlobal.fctFormatDateSave(oCompany, objFormGoodReceipt.Items.Item("9").Specific.string, 1)

                                            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                            StrSql = "SELECT T0.U_MISTPGRQ TotalGRFishQty , T0.U_MISESTSF EstimasiFish, T0.U_MISNFDIE MortalityQty, T0.U_MISTPCST TotalProjectCost, T0.U_MISTPGRC TotalGoodReceiptCost FROM " & _
                                                    "[@MIS_PRJMSTR] T0 " & _
                                                    "WHERE T0.U_MISPROID = '" & objColumns.Item("U_MISPROID").Cells.Item(1).Specific.string & "' "

                                            objRecSet.DoQuery(StrSql)

                                            If objRecSet.RecordCount > 0 Then

                                                For i As Integer = 1 To objMatrix.RowCount - 1
                                                    If objColumns.Item("1").Cells.Item(i).Specific.string <> "" Then
                                                        If objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.string = "" Then
                                                            objApplication.StatusBar.SetText("Fish Qty Must Fill ~OIGN.2.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            TotalFishQty = 0 + TotalFishQty
                                                            QtyInKg = 0 + QtyInKg
                                                            GoTo Setnothing
                                                        Else
                                                            TotalFishQty = objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.string + TotalFishQty
                                                            QtyInKg = objColumns.Item("9").Cells.Item(i).Specific.string + QtyInKg
                                                        End If
                                                    End If
                                                Next

                                                TotalGRFishQty = objRecSet.Fields.Item("TotalGRFishQty").Value
                                                EstimasiFish = objRecSet.Fields.Item("EstimasiFish").Value
                                                MortalityQty = objRecSet.Fields.Item("MortalityQty").Value
                                                TotalProjectCost = objRecSet.Fields.Item("TotalProjectCost").Value
                                                TotalGoodReceiptCost = objRecSet.Fields.Item("TotalGoodReceiptCost").Value
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

                                                Dim VarTotal As Double


                                                For i As Integer = 1 To objMatrix.RowCount - 1
                                                    objColumns.Item("10").Cells.Item(i).Specific.string = 0
                                                    If objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "N" Then
                                                        If (TotalFishQty + TotalGRFishQty) < (EstimasiFish - MortalityQty) Then
                                                            If objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.string = "" Then
                                                                objColumns.Item("14").Cells.Item(i).Specific.string = (0 / (EstimasiFish - MortalityQty)) * TotalProjectCost
                                                            Else
                                                                If TotalProjectCost <= TotalGoodReceiptCost Then
                                                                    objColumns.Item("10").Cells.Item(i).Specific.string = 0.01
                                                                    objFormGoodReceipt.Items.Item("11").Specific.string = "Value Kecil karena adanya Kesalahan input Last Harvesting"
                                                                Else
                                                                    objColumns.Item("14").Cells.Item(i).Specific.string = (objColumns.Item("U_MISFISHQ").Cells.Item(i).Specific.string / (EstimasiFish - MortalityQty)) * TotalProjectCost
                                                                End If
                                                            End If
                                                        ElseIf (TotalFishQty + TotalGRFishQty) > (EstimasiFish - MortalityQty) Then
                                                            objColumns.Item("10").Cells.Item(i).Specific.string = 1
                                                        End If

                                                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "Y" Then
                                                        If i = objMatrix.RowCount - 1 Then
                                                            If TotalProjectCost <= TotalGoodReceiptCost Then
                                                                objColumns.Item("10").Cells.Item(i).Specific.string = 0.01
                                                                objFormGoodReceipt.Items.Item("11").Specific.string = "Value Kecil karena adanya Kesalahan input Last Harvesting"
                                                            Else
                                                                objColumns.Item("14").Cells.Item(i).Specific.string = TotalProjectCost - TotalGoodReceiptCost - VarTotal
                                                            End If
                                                        Else
                                                            VarTotal = (objColumns.Item("9").Cells.Item(i).Specific.string / QtyInKg * (TotalProjectCost - TotalGoodReceiptCost)) + VarTotal
                                                            objColumns.Item("14").Cells.Item(i).Specific.string = (objColumns.Item("9").Cells.Item(i).Specific.string / QtyInKg * (TotalProjectCost - TotalGoodReceiptCost))
                                                        End If
                                                    End If
                                                Next

                                            Else
                                                objApplication.StatusBar.SetText("Project Master Not Available, Please Input Project Master ~OIGN.2.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                GoTo Setnothing
                                            End If
                                            'End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                            objMatrix = Nothing
                                            objColumns = Nothing
                                        End If

                                        'Calculate Total Processing Plant
                                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 3 Then
                                        Dim GiTotal As Double
                                        Dim GRTotalNotFresh As Double
                                        Dim GrTotalFresh As Double
                                        Dim GrQtyTotalFresh As Double

                                        GiTotal = objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Specific.value

                                        objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                        objColumns = objMatrix.Columns

                                        For i As Integer = 1 To objMatrix.RowCount - 1
                                            If objColumns.Item("1").Cells.Item(i).Specific.string <> "" Then
                                                If objColumns.Item("U_MISINFO").Cells.Item(i).Specific.value = "" Then
                                                    objApplication.StatusBar.SetText("You Must Fill Out Category First", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                                    GoTo Setnothing
                                                Else
                                                    If objColumns.Item("U_MISINFO").Cells.Item(i).Specific.value = 1 Then
                                                        GRTotalNotFresh = 0 + GRTotalNotFresh
                                                        GrQtyTotalFresh = objColumns.Item("9").Cells.Item(i).Specific.value + GrQtyTotalFresh
                                                    Else
                                                        GRTotalNotFresh = objColumns.Item("9").Cells.Item(i).Specific.value + GRTotalNotFresh
                                                    End If
                                                End If
                                            End If

                                        Next

                                        For i As Integer = 1 To objMatrix.RowCount - 1
                                            If objColumns.Item("1").Cells.Item(i).Specific.string <> "" Then
                                                If objColumns.Item("U_MISINFO").Cells.Item(i).Specific.value = "" Then
                                                    objApplication.StatusBar.SetText("You Must Fill Out Category First", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                                    GoTo Setnothing
                                                Else
                                                    If objColumns.Item("U_MISINFO").Cells.Item(i).Specific.value = 1 Then
                                                        GrTotalFresh = GiTotal - GRTotalNotFresh
                                                        objColumns.Item("14").Cells.Item(i).Specific.value = (objColumns.Item("9").Cells.Item(i).Specific.value / GrQtyTotalFresh) * GrTotalFresh
                                                    Else
                                                        objColumns.Item("14").Cells.Item(i).Specific.value = objColumns.Item("9").Cells.Item(i).Specific.value * 1
                                                    End If
                                                End If
                                            End If
                                        Next



                                    ElseIf objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = 4 Then
                                        Dim Yield As Double

                                        If objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string = "" Then
                                            objApplication.StatusBar.SetText("Refer Document Must Be Fill ~OIGN.2.0005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            GoTo Setnothing
                                        Else
                                            ReferDocument = objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string

                                            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                                            objColumns = objMatrix.Columns

                                            StrSql = "SELECT DISTINCT E4.Debit TotalGI FROM OIGE E0 INNER JOIN IGE1 E1 ON E0.DocEntry = E1.DocEntry " & _
                                                    "LEFT JOIN OJDT E3 ON E0.DocNum = E3.BaseRef AND E0.ObjType = E3.TransType INNER JOIN JDT1 E4 " & _
                                                    "ON E3.TransId = E4.TransId AND E1.AcctCode = E4.Account WHERE(E0.U_MISTRXTP = 6) AND E0.docnum = " & ReferDocument & " "

                                            objRecSet.DoQuery(StrSql)
                                            If objRecSet.RecordCount = 1 Then
                                                StrSql = "SELECT DISTINCT E4.Debit TotalGI FROM OIGE E0 INNER JOIN IGE1 E1 ON E0.DocEntry = E1.DocEntry " & _
                                                        "LEFT JOIN OJDT E3 ON E0.DocNum = E3.BaseRef AND E0.ObjType = E3.TransType INNER JOIN JDT1 E4 " & _
                                                        "ON E3.TransId = E4.TransId AND E1.AcctCode = E4.Account WHERE(E0.U_MISTRXTP = 6) AND E0.docnum = " & ReferDocument & " "
                                            Else
                                                StrSql = "SELECT SUM(T0.StockPrice * T0.Quantity) TotalGI FROM IGE1 T0 " & _
                                                            "INNER JOIN OIGE T1 ON T0.DocEntry = T1.DocEntry " & _
                                                            "Where T1.docnum = " & ReferDocument & " "
                                            End If

                                            TotalGI = objRecSet.Fields.Item("TotalGI").Value

                                            For i As Integer = 1 To objMatrix.RowCount - 1
                                                If objColumns.Item("1").Cells.Item(i).Specific.string <> "" Then
                                                    StrSql = "SELECT U_MISFGRDF Flag, U_MISYIELD Yield From OITM WHERE ItemCode = '" & objColumns.Item("1").Cells.Item(i).Specific.string & "' "
                                                    objRecSet.DoQuery(StrSql)
                                                    Flag = objRecSet.Fields.Item("Flag").Value

                                                    If Flag = "" Then
                                                        'objApplication.ActivateMenuItem(3073)

                                                        objApplication.MessageBox("U Must Fill First Grade In Item Master Data ~OIGN.2.0007~", 1, "OK")
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                                                        GoTo Setnothing
                                                    ElseIf Flag = "Y" Then
                                                        If objRecSet.Fields.Item("Yield").Value = 0 Then
                                                            'objApplication.ActivateMenuItem(3073)
                                                            objApplication.MessageBox("U Must Fill Yield In Item Master Data ~OIGN.2.0007~", 1, "OK")
                                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                                                            GoTo Setnothing
                                                        Else
                                                            Yield = objRecSet.Fields.Item("Yield").Value
                                                            objColumns.Item("U_MISYIELD").Cells.Item(i).Specific.string = Yield
                                                            objColumns.Item("U_MISNFPRO").Cells.Item(i).Specific.string = (objColumns.Item("9").Cells.Item(i).Specific.string / Yield) * 100
                                                            objColumns.Item("10").Cells.Item(i).Specific.string = 0
                                                            objColumns.Item("14").Cells.Item(i).Specific.string = 0
                                                            TotalGRNOTFirstGrade = 0 + TotalGRNOTFirstGrade
                                                            TotalConvertionFish = objColumns.Item("U_MISNFPRO").Cells.Item(i).Specific.string + TotalConvertionFish
                                                        End If
                                                    ElseIf Flag = "N" Then
                                                        objColumns.Item("U_MISNFPRO").Cells.Item(i).Specific.string = 0
                                                        objColumns.Item("10").Cells.Item(i).Specific.string = 1
                                                        TotalGRNOTFirstGrade = objColumns.Item("9").Cells.Item(i).Specific.string + TotalGRNOTFirstGrade
                                                        TotalConvertionFish = 0 + TotalConvertionFish
                                                    End If
                                                End If
                                            Next

                                            For i As Integer = 1 To objMatrix.RowCount - 1
                                                If objColumns.Item("1").Cells.Item(i).Specific.string <> "" Then
                                                    StrSql = "SELECT U_MISFGRDF Flag From OITM WHERE ItemCode = '" & objColumns.Item("1").Cells.Item(i).Specific.string & "' "
                                                    objRecSet.DoQuery(StrSql)
                                                    Flag = objRecSet.Fields.Item("Flag").Value
                                                    If Flag = "Y" Then
                                                        ConvertionFish = objColumns.Item("U_MISNFPRO").Cells.Item(i).Specific.string
                                                        LineTotal = ConvertionFish / TotalConvertionFish * (TotalGI - TotalGRNOTFirstGrade)
                                                        objColumns.Item("14").Cells.Item(i).Specific.string = LineTotal
                                                        'objColumns.Item("10").Cells.Item(i).Specific.string = (LineTotal / objColumns.Item("9").Cells.Item(i).Specific.string)
                                                    ElseIf Flag = "N" Then
                                                        LineTotal = objColumns.Item("9").Cells.Item(i).Specific.string
                                                        'objColumns.Item("10").Cells.Item(i).Specific.string = (LineTotal / objColumns.Item("9").Cells.Item(i).Specific.string)
                                                    End If
                                                End If
                                            Next

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                                            objMatrix = Nothing
                                            objColumns = Nothing
                                        End If
                                    Else
                                        objApplication.StatusBar.SetText("SuccessFull ~OIGN.2.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)

                                    End If
                                End If
                            End If
                    End Select

                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    '    If objFormGoodReceipt.Items.Item("1").Specific.caption = "ADD" Then
                    '        If pVal.ColUID = "U_MISFISHQ" Or pVal.ColUID = "9" Or pVal.ItemUID = "U_MISLASTB" Then

                    '        End If
                    '    End If
            End Select
        End If

        'List Good Issue
        If pVal.FormTypeEx = ListGI_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormListGI = objApplication.Forms.Item(pVal.FormUID)
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                blnModalListGI = False
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    'If fctFormExist(GoodIssue_FormId, intRowGoodIssueDetail) Then
                    '    objApplication.Forms.Item(intRowGoodIssueDetail).Select()
                    'Else
                    Dim DocEntry As Integer
                    If (pVal.ColUID = "DocEntry" Or pVal.ColUID = "Number" Or pVal.ColUID = "series" Or pVal.ColUID = "DocDate" Or pVal.ColUID = "RitNo") And (pVal.BeforeAction) And (pVal.Row > 0) Then
                        blnModalListGI = False
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing

                        objMatrix = objFormListGI.Items.Item("MtxGI").Specific
                        objColumns = objMatrix.Columns
                        DocEntry = objColumns.Item("DocEntry").Cells.Item(pVal.Row).Specific.string

                        subInsertDataIntoGoodReceipt(DocEntry, objFormGoodReceipt, pVal)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing
                    End If
            End Select
        End If

        ' UDF GoodReceipt
        If pVal.FormTypeEx = GoodReceiptUDF_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormGoodReceiptUDF = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ItemUID
                        Case "U_MISLASTB"
                            If Not pVal.BeforeAction Then
                                Dim TransactionType As Integer
                                Dim IntMsg As Integer

                                TransactionType = objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string
                                If objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "" Then

                                    If TransactionType = 1 Then
                                        IntMsg = objApplication.MessageBox("Are You Sure To Last Harvest? ", 1, "Yes", "No")
                                        If IntMsg = 1 Then
                                            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "Y"
                                        Else
                                            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Specific.string = "N"
                                        End If
                                    End If
                                Else
                                End If
                            End If

                        Case "U_MISTRXTP"
                            If Not pVal.BeforeAction Then
                                Dim TransactionType As Integer
                                If objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string = "" Then
                                    objApplication.StatusBar.SetText("Transaction Type Must Be Fill", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo Setnothing
                                Else
                                    TransactionType = objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string
                                    subLostFocusGR(TransactionType)
                                End If
                            End If
                    End Select
            End Select

        End If

        ' Project Harvest
        If pVal.FormTypeEx = ProjectHarvest_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormProjectHarvest = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "BtnNETID"
                            If pVal.BeforeAction Then
                                subFormLoadLookUp("SearchNetIdHarvest")
                            End If

                        Case "btnOK"
                            If pVal.BeforeAction Then
                                If UCase(oFormProjectHarvest.Items.Item(pVal.ItemUID).Specific.Caption) = "FIND" Then
                                    subFPFindData("Harvest", oFormProjectHarvest, BubbleEvent)
                                ElseIf UCase(oFormProjectHarvest.Items.Item(pVal.ItemUID).Specific.Caption) = "HARVEST" Then
                                    Dim ProjectId As String
                                    ProjectId = oFormProjectHarvest.Items.Item("MISPROID").Specific.string

                                    If oFormProjectHarvest.Items.Item("MISHARVD").Specific.string = "" Then
                                        objApplication.StatusBar.SetText("Actual Harvest Date Must Be Fill ~MISPRJHRVST.1.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        'ElseIf oFormProjectHarvest.Items.Item("MISHARVQ").Specific.string = "" Or oFormProjectHarvest.Items.Item("MISHARVQ").Specific.string = 0 Then
                                        '    objApplication.StatusBar.SetText("Actual Harvest Qty Must Be Fill", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        GoTo Setnothing
                                    End If

                                    subHarvest(ProjectId)
                                    'GoTo Setnothing
                                ElseIf UCase(oFormProjectHarvest.Items.Item(pVal.ItemUID).Specific.Caption) = "OK" Then
                                    oFormProjectHarvest.Close()
                                End If
                            End If
                    End Select

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ItemUID
                        Case "MISNETID"
                            Dim objrecSet As SAPbobsCOM.Recordset = Nothing
                            Dim strsql As String
                            Dim NetId As String
                            Dim Tanggal As String

                            objrecSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                            If oFormProjectHarvest.Items.Item("MISNETID").Specific.string <> "" Then
                                '    objApplication.StatusBar.SetText("You Must Fill Net ID", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                'Else
                                NetId = oFormProjectHarvest.Items.Item("MISNETID").Specific.string
                                strsql = "SELECT TOP 1 U_MISSIGND Tanggal FROM [@MIS_PRJMSTR] WHERE U_MISNETID = '" & NetId & "' AND U_MISNETST = 'O' "
                                objrecSet.DoQuery(strsql)

                                Tanggal = ClsGlobal.fctFormatDate(objrecSet.Fields.Item("Tanggal").Value, oCompany)
                                oFormProjectHarvest.Items.Item("MISSIGND").Specific.string = Tanggal
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objrecSet)


                        Case "MISSIGND"
                            Dim NetId As String
                            Dim StockDate As String
                            Dim StockdateValue As Date
                            Dim Tahun As Integer
                            Dim Bulan As Integer
                            Dim StrBulan As String
                            Dim Hari As Integer
                            Dim StrHari As String
                            NetId = oFormProjectHarvest.Items.Item("MISNETID").Specific.string
                            StockDate = oFormProjectHarvest.Items.Item("MISSIGND").Specific.string

                            If NetId <> "" And StockDate <> "" Then
                                If Len(Trim(NetId)) = 8 Then
                                    StockdateValue = CDate(ClsGlobal.fctFormatDateSave(oCompany, oFormProjectHarvest.Items.Item("MISSIGND").Specific.string, 5))
                                    Tahun = StockdateValue.Year
                                    Bulan = StockdateValue.Month
                                    If Len(Trim(Bulan)) = 1 Then
                                        StrBulan = "0" + CStr(Bulan)
                                    Else
                                        StrBulan = CStr(Bulan)
                                    End If
                                    Hari = StockdateValue.Day
                                    If Len(Trim(Hari)) = 1 Then
                                        StrHari = "0" + CStr(Hari)
                                    Else
                                        StrHari = CStr(Hari)
                                    End If
                                    oFormProjectHarvest.Items.Item("MISPROID").Specific.string = NetId + "." + CStr(Tahun) + StrBulan + StrHari
                                Else
                                    objApplication.MessageBox("Net Id Length must be 8 digit (FarmId + Net Type + No Farm Id) ~MISPRJHRVST.2.0001~", 1, "Ok")
                                End If
                            End If
                    End Select

                    'Select Case pVal.EventType
                    '    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    '        Select Case pVal.ItemUID

                    '            Case "BtnNETID"
                    '                If pVal.BeforeAction Then
                    '                    subFormLoadLookUp("SearchNetIdHarvest")
                    '                End If
                    '        End Select
            End Select

        End If

        If pVal.FormTypeEx = LookUpMortal_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormLookUpMortal = objApplication.Forms.Item(pVal.FormUID)
            End If
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "BtnADD"
                            'karno
                            If Not fctValidateMortal(oFormLookUpMortal, BubbleEvent) Then GoTo Setnothing

                    End Select

            End Select

        End If
        'Project Master

        If pVal.FormTypeEx = ProjectMaster_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                objFormProjectMaster = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID

                        Case "BtnMortq"

                            If pVal.BeforeAction Then
                                If objFormProjectMaster.Items.Item("btnOK").Specific.Caption = "UPDATE" Then
                                    subFormLoadLookUpMortal("SearchMortal")
                                End If
                            End If


                        Case "BtnNETID"
                            If pVal.BeforeAction Then
                                subFormLoadLookUp("SearchNetId")
                            End If

                        Case "btnOK"
                            If pVal.BeforeAction Then
                                If UCase(objFormProjectMaster.Items.Item(pVal.ItemUID).Specific.Caption) = "ADD" Then
                                    If Not fctValidate(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption), BubbleEvent) Then GoTo Setnothing
                                    If Not fctTestingSave(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption), BubbleEvent) Then GoTo Setnothing
                                    SubProjectMasterModeAdd()
                                    'GoTo Setnothing
                                ElseIf UCase(objFormProjectMaster.Items.Item(pVal.ItemUID).Specific.Caption) = "UPDATE" Then
                                    If Not fctValidate(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption), BubbleEvent) Then GoTo Setnothing
                                    If Not fctTestingSave(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption), BubbleEvent) Then GoTo Setnothing
                                    SubProjectMasterModeAdd()
                                    'GoTo Setnothing
                                ElseIf UCase(objFormProjectMaster.Items.Item(pVal.ItemUID).Specific.Caption) = "FIND" Then
                                    subFPFindData("Master", objFormProjectMaster, BubbleEvent)
                                    'GoTo Setnothing
                                ElseIf UCase(objFormProjectMaster.Items.Item(pVal.ItemUID).Specific.Caption) = "OK" Then
                                    objFormProjectMaster.Close()

                                End If
                            End If
                            'Case "BtnSCIES"
                            '    If pVal.BeforeAction Then
                            '        subFormLoadLookUpSpecies("SearchSpecies")
                            '    End If
                    End Select

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ItemUID
                        'Case "MISFEEDQ"
                        '    If objFormProjectMaster.Items.Item("btnOK").Specific.Caption <> "FIND" Then
                        '        Dim FeedQty As Double
                        '        Dim FeedRatio As Double
                        '        Dim FishEstimatedInKg As Double
                        '        Dim FeedEstimated As Double
                        '        Dim InitialFishQty As Double

                        '        FeedQty = objFormProjectMaster.Items.Item("MISFEEDQ").Specific.string
                        '        FeedRatio = objFormProjectMaster.Items.Item("MISFCR").Specific.string
                        '        FeedEstimated = (FeedQty * FeedRatio) / 100
                        '        FishEstimatedInKg = InitialFishQty + FeedEstimated
                        '        objFormProjectMaster.Items.Item("MISFCE").Specific.string = FeedEstimated
                        '        objFormProjectMaster.Items.Item("MISTEFQK").Specific.string = FishEstimatedInKg
                        '    End If

                        Case "MISSIGND"
                            If objFormProjectMaster.Items.Item("btnOK").Specific.Caption <> "FIND" Then
                                Dim NetId As String
                                Dim StockDate As String
                                Dim StockdateValue As Date
                                Dim Tahun As Integer
                                Dim Bulan As Integer
                                Dim StrBulan As String
                                Dim Hari As Integer
                                Dim StrHari As String

                                Dim CultureDay As Integer

                                NetId = objFormProjectMaster.Items.Item("MISNETID").Specific.string
                                StockDate = objFormProjectMaster.Items.Item("MISSIGND").Specific.string

                                If NetId <> "" And StockDate <> "" Then
                                    If Len(Trim(NetId)) = 8 Then
                                        StockdateValue = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISSIGND").Specific.string, 5))
                                        Tahun = StockdateValue.Year
                                        Bulan = StockdateValue.Month
                                        If Len(Trim(Bulan)) = 1 Then
                                            StrBulan = "0" + CStr(Bulan)
                                        Else
                                            StrBulan = CStr(Bulan)
                                        End If
                                        Hari = StockdateValue.Day
                                        If Len(Trim(Hari)) = 1 Then
                                            StrHari = "0" + CStr(Hari)
                                        Else
                                            StrHari = CStr(Hari)
                                        End If
                                        objFormProjectMaster.Items.Item("MISPROID").Specific.string = NetId + "." + CStr(Tahun) + StrBulan + StrHari
                                        If objFormProjectMaster.Items.Item("MISHARVP").Specific.string = "" Then
                                            objFormProjectMaster.Items.Item("MISHARVP").Specific.string = 0
                                            objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ""
                                        End If
                                        If objFormProjectMaster.Items.Item("MISSIGND").Specific.string <> "" And objFormProjectMaster.Items.Item("MISHARVP").Specific.string <> 0 Then
                                            'StockDate = objFormProjectMaster.Items.Item("MISSIGND").Specific.string
                                            CultureDay = objFormProjectMaster.Items.Item("MISHARVP").Specific.string
                                            'StockdateValue = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISSIGND").Specific.string, 5))
                                            objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ClsGlobal.fctFormatDate(DateAdd(DateInterval.Day, CultureDay, StockdateValue), oCompany)
                                            'Else
                                            '    objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ""
                                        End If

                                    Else
                                        objApplication.MessageBox("Net Id Length must be 8 digit (FarmId + Net Type + No Farm Id) ~MISPRJMSTR.1.0001~", 1, "Ok")
                                    End If
                                End If

                            End If

                            'Case "MISHARVP"
                            '    If objFormProjectMaster.Items.Item("btnOK").Specific.Caption <> "FIND" Then
                            '        Dim StockDate As String
                            '        Dim CultureDay As Integer
                            '        Dim StockdateValue As Date

                            '        If objFormProjectMaster.Items.Item("MISHARVP").Specific.string = "" Then
                            '            objFormProjectMaster.Items.Item("MISHARVP").Specific.string = 0
                            '            objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ""
                            '        End If
                            '        If objFormProjectMaster.Items.Item("MISSIGND").Specific.string <> "" And objFormProjectMaster.Items.Item("MISHARVP").Specific.string <> 0 Then
                            '            StockDate = objFormProjectMaster.Items.Item("MISSIGND").Specific.string
                            '            CultureDay = objFormProjectMaster.Items.Item("MISHARVP").Specific.string
                            '            StockdateValue = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISSIGND").Specific.string, 5))
                            '            objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ClsGlobal.fctFormatDate(DateAdd(DateInterval.Day, CultureDay, StockdateValue), oCompany)
                            '            GoTo setnothing
                            '            'Else
                            '            '    objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ""
                            '        End If
                            '    End If

                        Case "MISNETID", "MISSCIES", "MISHATGO"

                            If objFormProjectMaster.Items.Item("btnOK").Specific.Caption = "ADD" Then
                                Dim NetId As String
                                Dim FarmCode As String
                                Dim Species As String
                                Dim Klasifikasi As String

                                If objFormProjectMaster.Items.Item("MISNETID").Specific.string = "" Then
                                    objFormProjectMaster.Items.Item("MISNETPUCD").Specific.string = ""
                                ElseIf objFormProjectMaster.Items.Item("MISHATGO").Specific.string <> "" And objFormProjectMaster.Items.Item("MISNETID").Specific.string <> "" And objFormProjectMaster.Items.Item("MISSCIES").Specific.string <> "" Then
                                    NetId = objFormProjectMaster.Items.Item("MISNETID").Specific.string
                                    FarmCode = Left(Trim(NetId), 3)
                                    Species = objFormProjectMaster.Items.Item("MISSCIES").Specific.string
                                    Klasifikasi = objFormProjectMaster.Items.Item("MISHATGO").Specific.string
                                    subValidateFarmcode(FarmCode, Species, Klasifikasi)
                                    'GoTo setnothing
                                    'Else
                                    '    GoTo Setnothing
                                End If

                            End If

                            'Case "MISSCIES"
                            '    If objFormProjectMaster.Items.Item("btnOK").Specific.Caption <> "FIND" Then
                            '        Dim Species As String
                            '        Dim NetCode As String
                            '        Dim strSql As String
                            '        Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                            '        Species = objFormProjectMaster.Items.Item("MISSCIES").Specific.string
                            '        NetCode = objFormProjectMaster.Items.Item("MISNETID").Specific.string
                            '        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            '        If objFormProjectMaster.Items.Item("MISSCIES").Specific.string = "" Then
                            '            objFormProjectMaster.Items.Item("MISESTLF").Specific.string = 100
                            '        ElseIf Species <> "" And NetCode <> "" Then
                            '            strSql = "select U_MISSURVR SurvivalRate from [@MIS_FSURVR] WHERE U_MISFARMC = '" & Left(NetCode, 3) & "' AND U_MISFSPEC = '" & Species & "' "
                            '            objRecSet.DoQuery(strSql)

                            '            If objRecSet.RecordCount > 0 Then
                            '                objFormProjectMaster.Items.Item("MISESTLF").Specific.string = objRecSet.Fields.Item("SurvivalRate").Value
                            '            Else
                            '                objFormProjectMaster.Items.Item("MISESTLF").Specific.string = 100
                            '            End If
                            '        End If

                            '        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

                            '    End If

                        Case "MISESTSF"
                            If objFormProjectMaster.Items.Item("btnOK").Specific.Caption <> "FIND" Then
                                Dim Estimated As Integer
                                Dim Rate As Double
                                Dim EstimatedQty As Integer

                                If objFormProjectMaster.Items.Item("MISESTSF").Specific.value = "" Then
                                    objFormProjectMaster.Items.Item("MISESTSF").Specific.value = 0
                                    objFormProjectMaster.Items.Item("MISESTHQ").Specific.value = ""
                                End If

                                If objFormProjectMaster.Items.Item("MISESTSF").Specific.value <> 0 And objFormProjectMaster.Items.Item("MISESTLF").Specific.value <> 0 Then
                                    Estimated = objFormProjectMaster.Items.Item("MISESTSF").Specific.value
                                    Rate = objFormProjectMaster.Items.Item("MISESTLF").Specific.value
                                    EstimatedQty = (Estimated * Rate) / 100

                                    objFormProjectMaster.Items.Item("MISESTHQ").Specific.value = EstimatedQty
                                    'Else
                                    '    objFormProjectMaster.Items.Item("MISESTSF").Specific.string = 0
                                    '    objFormProjectMaster.Items.Item("MISESTHD").Specific.string = ""
                                End If
                            End If


                    End Select

            End Select

        End If

        If pVal.FormTypeEx = LookUpNet_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormLookUpNet = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If (pVal.ColUID = "colNetCd" Or pVal.ColUID = "colFarmCd" Or pVal.ColUID = "colReg" Or pVal.ColUID = "colNetCap" Or pVal.ColUID = "colNetPu" Or pVal.ColUID = "ColNetSts") And pVal.BeforeAction And (pVal.Row > 0) Then
                        blnModalLookUp = False
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing
                        objMatrix = oFormLookUpNet.Items.Item("mtxSearch").Specific
                        objColumns = objMatrix.Columns

                        oFormProjectHarvest.Items.Item("MISNETID").Specific.String = objColumns.Item("colNetCd").Cells.Item(pVal.Row).Specific.String
                        oFormProjectHarvest.Items.Item("MISNETPUCD").Specific.String = objColumns.Item("colNetPu").Cells.Item(pVal.Row).Specific.String

                        BubbleEvent = False

                        oFormLookUpNet.Close()

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing
                    End If
            End Select
        End If


        If pVal.FormTypeEx = LookUp_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormLookUp = objApplication.Forms.Item(pVal.FormUID)
            End If

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If (pVal.ColUID = "colNetCd" Or pVal.ColUID = "colFarmCd" Or pVal.ColUID = "colReg" Or pVal.ColUID = "colNetCap" Or pVal.ColUID = "colNetPu" Or pVal.ColUID = "ColNetSts") And pVal.BeforeAction And (pVal.Row > 0) Then
                        blnModalLookUp = False
                        'Dim strSQL As String
                        'Dim FarmCode As String
                        '        Dim objrecSet As SAPbobsCOM.Recordset = Nothing
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing
                        objMatrix = oFormLookUp.Items.Item("mtxSearch").Specific
                        objColumns = objMatrix.Columns

                        objFormProjectMaster.Items.Item("MISNETID").Specific.String = ""
                        objFormProjectMaster.Items.Item("MISNETID").Specific.String = objColumns.Item("colNetCd").Cells.Item(pVal.Row).Specific.String

                        'FarmCode = objColumns.Item("colFarmCd").Cells.Item(pVal.Row).Specific.String

                        objFormProjectMaster.Items.Item("MISNETPUCD").Specific.String = ""
                        objFormProjectMaster.Items.Item("MISNETPUCD").Specific.String = objColumns.Item("colNetPu").Cells.Item(pVal.Row).Specific.String

                        '        ' Jika UDF Species, clasifikasi Berdasarkan pada Net Master
                        '        strSQL = "SELECT U_MISHARVP DayOfCulture, T0.U_MISHATGRO Clasification, " & _
                        '                 "T0.U_MISFSPEC Species, U_MISSURVR NetRate, U_MISFCR NetFCR ,T0.U_MISFARMC " & _
                        '                 "FROM [@MIS_RATE] T0 " & _
                        '                 "INNER JOIN [@MIS_NETMS] T1 " & _
                        '                 "ON T0.U_MISFARMC = T1.U_MISFARMC " & _
                        '                 "AND T0.U_MISFSPEC = T1.U_MISFSPEC " & _
                        '                 "AND T0.U_MISHATGRO = T1.U_MISHATGRO WHERE U_MISFARMC = '" & FarmCode & "' "


                        '        '' Jika UDF Species, clasifikasi Berdasarkan pada Farm Master
                        '        'strSQL = "SELECT U_MISHARVP DayOfCulture, T0.U_MISHATGRO Clasification, " & _
                        '        '         "T0.U_MISFSPEC Species, U_MISSURVR NetRate, U_MISFCR NetFCR ,T0.U_MISFARMC " & _
                        '        '         "FROM [@MIS_RATE] T0 " & _
                        '        '         "INNER JOIN [@MIS_NETMS] T1 " & _
                        '        '         "ON T0.U_MISFARMC = T1.U_MISFARMC " & _
                        '        '         "AND T0.U_MISFSPEC = T1.U_MISFSPEC " & _
                        '        '         "AND T0.U_MISHATGRO = T1.U_MISHATGRO WHERE U_MISFARMC = '" & FarmCode & "' "


                        '        objrecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '        objrecSet.DoQuery(strSQL)

                        '        If objrecSet.RecordCount > 0 Then
                        '            objFormProjectMaster.Items.Item("MISESTLF").Specific.String = 0
                        '            objFormProjectMaster.Items.Item("MISESTLF").Specific.String = objrecSet.Fields.Item("NetRate").Value
                        '            objFormProjectMaster.Items.Item("MISFCR").Specific.String = 0
                        '            objFormProjectMaster.Items.Item("MISFCR").Specific.String = objrecSet.Fields.Item("NetFCR").Value
                        '            objFormProjectMaster.Items.Item("MISHATGO").Specific.String = ""
                        '            objFormProjectMaster.Items.Item("MISHATGO").Specific.String = objrecSet.Fields.Item("Clasification").Value
                        '            objFormProjectMaster.Items.Item("MISSCIES").Specific.String = ""
                        '            objFormProjectMaster.Items.Item("MISSCIES").Specific.String = objrecSet.Fields.Item("Species").Value
                        '            objFormProjectMaster.Items.Item("MISHARVP").Specific.String = 0
                        '            objFormProjectMaster.Items.Item("MISHARVP").Specific.String = objrecSet.Fields.Item("DayOfCulture").Value

                        '        Else
                        '            objApplication.StatusBar.SetText("U Must Fill Rate Master", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '            GoTo Setnothing
                        '        End If
                        oFormLookUp.Close()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        '        System.Runtime.InteropServices.Marshal.ReleaseComObject(objrecSet)
                        objMatrix = Nothing
                        objColumns = Nothing
                        '        objrecSet = Nothing
                    End If
            End Select
        End If

        If pVal.FormTypeEx = LookUpBatch_FormId Then
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                oFormLookUpBatch = objApplication.Forms.Item(pVal.FormUID)
            End If
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If (pVal.ColUID = "colFinGB" Or pVal.ColUID = "colStrain" Or pVal.ColUID = "colAcro" Or pVal.ColUID = "colDesc") And pVal.BeforeAction And (pVal.Row > 0) Then
                        blnModalLookUp = False
                        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim objColumns As SAPbouiCOM.Columns = Nothing
                        objMatrix = oFormLookUpBatch.Items.Item("mtxSrcBtch").Specific
                        objColumns = objMatrix.Columns

                        objFormProjectMaster.Items.Item("MISGENCD").Specific.String = objColumns.Item("colFinGB").Cells.Item(pVal.Row).Specific.String
                        BubbleEvent = False

                        oFormLookUpBatch.Close()

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
                        objMatrix = Nothing
                        objColumns = Nothing
                    End If
            End Select
        End If


Setnothing:
        '        objMatrix = Nothing
        '        objColumns = Nothing
    End Sub

    Private Sub subHarvest(ByVal Project As String)
        Dim intMsg As Integer
        Dim strSql As String
        Dim ActualHarvestQty As Double
        Dim ActualHarvestDate As Date
        Dim InitFishQty As Double
        Dim FeedConsQty As Double
        Dim FeedConsRate As Double
        Dim FeedConsEstimate As Double
        Dim TotalFishEstimate As Double
        Dim InitFishCost As Double
        Dim FeedConsCost As Double
        Dim CumFishDie As Double
        Dim TotalEstimated As Double
        Dim TotalGRCost As Double
        Dim TotalGRQty As Double
        Dim ProjectCalc As String
        Dim Flag As Integer
        Dim ProjectHarvestRemarks As String
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing

        InitFishQty = oFormProjectHarvest.Items.Item("MISINIFQ").Specific.value
        FeedConsQty = oFormProjectHarvest.Items.Item("MISFEEDQ").Specific.value
        InitFishCost = oFormProjectHarvest.Items.Item("MISINIFC").Specific.value
        FeedConsCost = oFormProjectHarvest.Items.Item("MISFEEDC").Specific.value
        FeedConsRate = oFormProjectHarvest.Items.Item("MISFCR").Specific.value
        FeedConsEstimate = oFormProjectHarvest.Items.Item("MISFCE").Specific.value
        TotalFishEstimate = oFormProjectHarvest.Items.Item("MISTEFQK").Specific.value
        CumFishDie = oFormProjectHarvest.Items.Item("MISNFDIE").Specific.value
        TotalEstimated = oFormProjectHarvest.Items.Item("MISTPCST").Specific.value
        TotalGRCost = oFormProjectHarvest.Items.Item("MISTPGRC").Specific.value
        TotalGRQty = oFormProjectHarvest.Items.Item("MISTPGRQ").Specific.value
        ProjectCalc = oFormProjectHarvest.Items.Item("MISPROCS").Specific.string
        ActualHarvestQty = oFormProjectHarvest.Items.Item("MISHARVQ").Specific.string
        Flag = oFormProjectHarvest.Items.Item("MISPROCS").Specific.string
        ProjectHarvestRemarks = oFormProjectHarvest.Items.Item("MISPROHR").Specific.string

        If oFormProjectHarvest.Items.Item("MISHARVD").Specific.String = "" Then
            ActualHarvestDate = "12:00:00 AM"
        Else
            ActualHarvestDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, oFormProjectHarvest.Items.Item("MISHARVD").Specific.String, 3))
        End If



        intMsg = objApplication.MessageBox("Are You Sure To Harvest With Project Id = '" & Project & "'", 1, "YES", "NO")

        If intMsg = 1 Then
            strSql = "UPDATE [@MIS_PRJMSTR] SET U_MISINIFQ = " & InitFishQty & " , U_MISPROHR = '" & ProjectHarvestRemarks & "' " & _
            ", U_MISFEEDQ = " & FeedConsQty & " " & _
            ",U_MISINIFC = " & InitFishCost & " " & _
            ",U_MISFEEDC = " & FeedConsCost & ",U_MISFCR = " & FeedConsRate & ",U_MISFCE = " & FeedConsEstimate & ",U_MISTFQKG = " & TotalFishEstimate & " " & _
            ",U_MISNFDIE = " & CumFishDie & " " & _
            ",U_MISTPCST = " & TotalEstimated & "  " & _
            ",U_MISTPGRC = " & TotalGRCost & "  " & _
            ",U_MISTPGRQ = " & TotalGRQty & "  " & _
            ", U_MISNETST = 'H', U_MISHARVD = " & IIf(ActualHarvestDate = "12:00:00 AM", "NULL", "'" & ActualHarvestDate & "'") & " " & _
            ", U_MISHARVQ = " & ActualHarvestQty & " " & _
            ", U_MISPROCS = " & Flag & " " & _
            "FROM [@MIS_PRJMSTR] WHERE U_MISPROID = '" & Project & "' AND U_MISNETST <> 'D' "

            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecSet.DoQuery(strSql)

            oFormProjectHarvest.Items.Item("btnOK").Specific.caption = "OK"
            oFormProjectHarvest.Items.Item("MISNETST").Specific.string = "H"

            objApplication.StatusBar.SetText("Update Successfully ~12.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        Else
            GoTo setnothing
        End If



setnothing:
        objRecSet = Nothing


    End Sub

    Private Sub subGeneratedBatchFG()
        Dim ObjRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim Region As String
        Dim Tanggal As Date
        Dim RunningNo As String
        Dim Batch As String
        Dim Tahun As Integer
        Dim FullTanggal As String
        Dim strSql As String
        Dim intMsg As Integer

        Region = oFormBatchFG.Items.Item("Region").Specific.string
        Tanggal = CDate(ClsGlobal.fctFormatDateSave(oCompany, oFormBatchFG.Items.Item("DocDate").Specific.string, 1))

        Tahun = Tanggal.Year

        FullTanggal = CStr(Right(Tahun, 2))

        ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strSql = "SELECT ISNULL(MAX(RIGHT(T0.[DistNumber],7)),0) RunningNo " & _
                "FROM OBTN T0  INNER JOIN OITM T1 " & _
                "ON T0.ItemCode = T1.ItemCode " & _
                "WHERE T1.[ItmsGrpCod] = 107 " & _
                "AND LEFT(T0.[DistNumber],1) = '" & Region & "'" & _
                "AND SUBSTRING(T0.[DistNumber],3,2) = '" & FullTanggal & "' "

        ObjRecSet.DoQuery(strSql)

        If ObjRecSet.RecordCount > 0 Then
            If ObjRecSet.Fields.Item("RunningNo").Value = 0 Then
                RunningNo = "0000001"
            Else
                RunningNo = ObjRecSet.Fields.Item("RunningNo").Value
                RunningNo = RunningNo + 1

                If Len(RunningNo) = 1 Then
                    RunningNo = "000000" + RunningNo
                ElseIf Len(RunningNo) = 2 Then
                    RunningNo = "00000" + RunningNo
                ElseIf Len(RunningNo) = 3 Then
                    RunningNo = "0000" + RunningNo
                ElseIf Len(RunningNo) = 4 Then
                    RunningNo = "000" + RunningNo
                ElseIf Len(RunningNo) = 5 Then
                    RunningNo = "00" + RunningNo
                ElseIf Len(RunningNo) = 6 Then
                    RunningNo = "0" + RunningNo
                ElseIf Len(RunningNo) = 7 Then
                    RunningNo = RunningNo
                End If

            End If

            Batch = Region + "." + FullTanggal + "." + RunningNo

            intMsg = objApplication.MessageBox("Are You Sure To Generate This Batch '" & Batch & "'", 1, "OK", "Cancel")

            If intMsg = 1 Then

                objMatrix = objFormBatch.Items.Item("3").Specific
                objColumns = objMatrix.Columns

                objColumns.Item("2").Cells.Item(1).Specific.string = Batch

                oFormBatchFG.Close()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
            Else
                oFormBatchFG.Close()
            End If

        End If


        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)
    End Sub

    Private Sub subGeneratedBatchRM()
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim NetCode As String
        Dim Tanggal As Date
        Dim BoxNo As String
        Dim RitNo As String
        Dim Batch As String
        Dim Tahun As Integer
        Dim bulan As Integer
        Dim hari As Integer
        Dim StrBulan As String
        Dim StrHari As String
        Dim FullTanggal As String


        NetCode = oFormBatchRM.Items.Item("NetCd").Specific.string
        Tanggal = CDate(ClsGlobal.fctFormatDateSave(oCompany, oFormBatchRM.Items.Item("DocDate").Specific.string, 1))

        Tahun = Tanggal.Year
        bulan = Tanggal.Month
        If Len(Trim(bulan)) = 1 Then
            StrBulan = "0" + CStr(bulan)
        Else
            StrBulan = CStr(bulan)
        End If
        hari = Tanggal.Day
        If Len(Trim(hari)) = 1 Then
            StrHari = "0" + CStr(hari)
        Else
            StrHari = CStr(hari)
        End If

        FullTanggal = CStr(Right(Tahun, 2)) + StrBulan + StrHari

        RitNo = oFormBatchRM.Items.Item("RitNo").Specific.string
        BoxNo = oFormBatchRM.Items.Item("BoxNo").Specific.string

        Batch = NetCode + "." + FullTanggal + "." + RitNo + "." + BoxNo

        objMatrix = objFormBatch.Items.Item("3").Specific
        objColumns = objMatrix.Columns

        objColumns.Item("2").Cells.Item(1).Specific.string = Batch

        oFormBatchRM.Close()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)


    End Sub

    Private Sub subGeneratedItemGroup()
        Dim strSql As String
        Dim ObjRecSet As SAPbobsCOM.Recordset

        ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim RunningNumber As String
        Dim intNumber As Integer
        Dim ItemCodeFG As String
        Dim ItemCodeUF As String

        Dim ItemGroupDesc As String
        Dim Brand As String
        Dim BrandDesc As String
        Dim Grade As String
        Dim GradeDesc As String
        Dim Species As String
        Dim SpeciesDesc As String
        Dim SkinningNCut As String
        Dim SkinningNCutDesc As String
        Dim Sizing As String
        Dim SizingDesc As String
        Dim TreatmentGlaz As String
        Dim TreatmentGlazDesc As String
        Dim Condition As String
        Dim ConditionDesc As String
        Dim Bagging As String
        Dim BaggingDesc As String
        Dim intMsg As Integer
        Dim ObjItemGroup As SAPbouiCOM.ComboBox = Nothing
        Dim ItemGroup As String
        Dim ItemGroupDescMaster As String
        Dim CustomerGroup As String
        'Dim CustomerGroupDesc As String
        Dim NetWeight As String
        Dim NetWeightUnit As String
        ObjItemGroup = objFormItemMaster.Items.Item("39").Specific
        ItemGroup = ObjItemGroup.Selected.Value
        ItemGroupDescMaster = ObjItemGroup.Selected.Description


        If ItemGroup = "107" Or ItemGroup = "105" Or ItemGroupDescMaster = "Unpack Finish Goods" Or ItemGroupDescMaster = "Finish Goods" Then
            If objFormItemMasterUDF.Items.Item("U_MISSPECS").Specific.value = "" Then
                objApplication.StatusBar.SetText("Species Must Fill ~11.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                Species = Trim(objFormItemMasterUDF.Items.Item("U_MISSPECS").Specific.value)

                strSql = "Select Name Descript from [@MIS_SPEC] WHERE Code = '" & Species & "' "
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount <> 0 Then
                    SpeciesDesc = Trim(ObjRecSet.Fields.Item("Descript").Value)
                Else
                    objApplication.StatusBar.SetText("Species Name Not Found Check Master Species ~11.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                End If

            End If

            If objFormItemMasterUDF.Items.Item("U_MISCUTSK").Specific.value = "" Then
                objApplication.StatusBar.SetText("Cutting & Skinning Must Fill ~11.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                SkinningNCut = Trim(objFormItemMasterUDF.Items.Item("U_MISCUTSK").Specific.value)

                strSql = "Select Name Descript from [@MIS_SKINCUT] WHERE Code = '" & SkinningNCut & "' "
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount <> 0 Then
                    SkinningNCutDesc = ObjRecSet.Fields.Item("Descript").Value
                Else
                    objApplication.StatusBar.SetText("Cutting & Skinning Name Not Found Check Master Skinning ~11.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo setnothing
                End If

            End If

            'If objFormItemMasterUDF.Items.Item("U_MISCUTTI").Specific.selected.value = "" Then
            '    objApplication.StatusBar.SetText("Cutting Must Fill ~110005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '    GoTo Setnothing
            'Else
            '    Cutting = objFormItemMasterUDF.Items.Item("U_MISCUTTI").Specific.selected.value

            '    strSql = "Select Name Descript from [@MIS_CUT] WHERE Code = '" & Cutting & "' "
            '    ObjRecSet.DoQuery(strSql)

            '    If ObjRecSet.RecordCount <> 0 Then
            '        CuttingDesc = ObjRecSet.Fields.Item("Descript").Value
            '    Else
            '        objApplication.StatusBar.SetText("Cutting Name Not Found Check Master Cutting ~110006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '    End If

            'End If

            If objFormItemMasterUDF.Items.Item("U_MISSIZET").Specific.string = "" Then
                objApplication.StatusBar.SetText("Size Tag Must Fill ~11.0007~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                Sizing = objFormItemMasterUDF.Items.Item("U_MISSIZET").Specific.string

                'strSql = "Select Name Descript from [@MIS_SIZET] WHERE Code = '" & Sizing & "' "
                'ObjRecSet.DoQuery(strSql)

                'If ObjRecSet.RecordCount <> 0 Then
                '    SizingDesc = ObjRecSet.Fields.Item("Descript").Value
                'Else
                '    objApplication.StatusBar.SetText("Size Tag Name Not Found Check Master Size Tag ~11.0008~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'End If

            End If

            If objFormItemMasterUDF.Items.Item("U_MISTRGL").Specific.value = "" Then
                objApplication.StatusBar.SetText("Treatment & Glazing Must Fill ~11.0009~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                TreatmentGlaz = objFormItemMasterUDF.Items.Item("U_MISTRGL").Specific.value

                strSql = "Select Name Descript from [@MIS_TREATGLAZ] WHERE Code = '" & TreatmentGlaz & "' "
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount <> 0 Then
                    TreatmentGlazDesc = ObjRecSet.Fields.Item("Descript").Value
                Else
                    objApplication.StatusBar.SetText("Treatment & Glazing Name Not Found Check Master Treatment ~11.0010~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo setnothing
                End If

            End If

            If objFormItemMasterUDF.Items.Item("U_MISYIELD").Specific.string = 0 Then
                objApplication.StatusBar.SetText("Production Std Yield Must Fill ~11.0011~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormItemMasterUDF.Items.Item("U_MISFGRDF").Specific.value = "" Then
                objApplication.StatusBar.SetText("First Grade Flag Only Must Fill ~11.0012~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            ElseIf objFormItemMasterUDF.Items.Item("U_MISFGRDF").Specific.selected.value <> "Y" And objFormItemMasterUDF.Items.Item("U_MISFGRDF").Specific.selected.value <> "N" Then
                objApplication.StatusBar.SetText("First Grade Flag Only Must Y Or N ~11.0013~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If


            If objFormItemMasterUDF.Items.Item("U_MISBAGGI").Specific.value = "" Then
                objApplication.StatusBar.SetText("Bagging Must Fill ~11.0016~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                Bagging = objFormItemMasterUDF.Items.Item("U_MISBAGGI").Specific.selected.value

                strSql = "Select Name Descript from [@MIS_BAG] WHERE Code = '" & Bagging & "' "
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount <> 0 Then
                    BaggingDesc = ObjRecSet.Fields.Item("Descript").Value
                Else
                    objApplication.StatusBar.SetText("Bagging Name Not Found Check Master Bagging ~11.0017~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo setnothing
                End If

            End If

            If objFormItemMasterUDF.Items.Item("U_MISGRADE").Specific.value = "" Then
                objApplication.StatusBar.SetText("Grade Must Fill ~11.0018~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                GoTo Setnothing
            Else
                Grade = objFormItemMasterUDF.Items.Item("U_MISGRADE").Specific.selected.value

                strSql = "Select Name Descript from [@MIS_GRADE] WHERE Code = '" & Grade & "' "
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount <> 0 Then
                    GradeDesc = ObjRecSet.Fields.Item("Descript").Value
                Else
                    objApplication.StatusBar.SetText("Grade Name Not Found Check Master Grade ~11.0019~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo setnothing
                End If

            End If
            If ItemGroup = "107" Or ItemGroupDescMaster = "Finish Goods" Then
                ItemGroupDesc = "FG"

                If objFormItemMasterUDF.Items.Item("U_MISCONDI").Specific.value = "" Then
                    objApplication.StatusBar.SetText("Condition Must Fill ~11.0014~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                Else
                    Condition = objFormItemMasterUDF.Items.Item("U_MISCONDI").Specific.selected.value

                    strSql = "Select Name Descript from [@MIS_COND] WHERE Code = '" & Condition & "' "
                    ObjRecSet.DoQuery(strSql)

                    If ObjRecSet.RecordCount <> 0 Then
                        ConditionDesc = ObjRecSet.Fields.Item("Descript").Value
                    Else
                        objApplication.StatusBar.SetText("Condition Name Not Found Check Master Condition ~11.0015~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        GoTo setnothing
                    End If

                End If

                If objFormItemMasterUDF.Items.Item("U_MISNETWG").Specific.String = "" Then
                    objApplication.StatusBar.SetText("Net Weight Must Fill ~11.0020~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objFormItemMasterUDF.Items.Item("U_MISNETWG").Click()
                    GoTo Setnothing
                Else
                    NetWeight = objFormItemMasterUDF.Items.Item("U_MISNETWG").Specific.String
                End If

                If objFormItemMasterUDF.Items.Item("U_MISNETWU").Specific.String = "" Then
                    objApplication.StatusBar.SetText("Net Weight Unit Must Fill ~11.0020~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objFormItemMasterUDF.Items.Item("U_MISNETWU").Click()
                    GoTo Setnothing
                Else
                    NetWeightUnit = objFormItemMasterUDF.Items.Item("U_MISNETWU").Specific.String
                End If

                If objFormItemMasterUDF.Items.Item("U_MISBRAND").Specific.value = "" Then
                    objApplication.StatusBar.SetText("Brand Must Fill ~11.0020~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                Else
                    Brand = objFormItemMasterUDF.Items.Item("U_MISBRAND").Specific.selected.value
                    strSql = "Select Name Descript from [@MIS_BRAND] WHERE Code = '" & Brand & "' "
                    ObjRecSet.DoQuery(strSql)

                    If ObjRecSet.RecordCount <> 0 Then
                        BrandDesc = ObjRecSet.Fields.Item("Descript").Value
                    Else
                        objApplication.StatusBar.SetText("Brand Name Not Found Check Master Brand ~11.0021~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        GoTo setnothing
                    End If

                End If
                If objFormItemMasterUDF.Items.Item("U_MISCARDG").Specific.string = "" Then
                    objApplication.StatusBar.SetText("Customer Group Must Fill ~11.0020~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                Else
                    CustomerGroup = objFormItemMasterUDF.Items.Item("U_MISCARDG").Specific.string
                    '    strSql = "Select ItmsGrpNam from OITB WHERE itmsGrpCod = '" & CustomerGroup & "' "
                    '    ObjRecSet.DoQuery(strSql)

                    '    If ObjRecSet.RecordCount <> 0 Then
                    '        CustomerGroupDesc = ObjRecSet.Fields.Item("Descript").Value
                    '    Else
                    '        objApplication.StatusBar.SetText("Customer Group Name Not Found Check Business Partner ~11.0021~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '    End If

                End If

                ItemCodeFG = ItemGroupDesc + Brand + CustomerGroup + Species + Condition + SkinningNCut + Grade + Bagging

                strSql = "Select ISNULL(MAX(RIGHT(ItemCode,4)),0) RunNumber From OITM where ItmsGrpCod = '" & ItemGroup & "' AND ItemCode Like '" & ItemCodeFG & "%'"
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount > 0 Then
                    intNumber = ObjRecSet.Fields.Item("RunNumber").Value
                    Select Case Len(Trim(intNumber))
                        Case 1
                            RunningNumber = "000" + CStr(Trim(intNumber + 1))
                        Case 2
                            RunningNumber = "00" + CStr(Trim(intNumber + 1))
                        Case 3
                            RunningNumber = "0" + CStr(Trim(intNumber + 1))
                            'Case 4
                            'RunningNumber = "0" + CStr(Trim(intNumber + 1))
                        Case 4
                            RunningNumber = CStr(Trim(intNumber + 1))
                    End Select
                Else
                    RunningNumber = "0001"
                End If


                intMsg = objApplication.MessageBox("Are You Sure To Generated ? Item Code =  " & ItemCodeFG + RunningNumber & " AND Description = " & ConditionDesc + " " + SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc + " " + NetWeight + " " + NetWeightUnit + " " + BrandDesc & "", 1, "Yes", "No", "")

                If intMsg = 1 Then
                    If Len(ItemCodeFG + RunningNumber) <= 20 Then
                        objFormItemMaster.Items.Item("5").Specific.string = ItemCodeFG + RunningNumber
                    Else
                        objFormItemMaster.Items.Item("5").Specific.string = Left(ItemCodeFG + RunningNumber, 20)
                        objApplication.StatusBar.SetText("Max Item code 20 Digit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                    If Len(ConditionDesc + " " + SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc + " " + NetWeight + " " + NetWeightUnit + " " + BrandDesc) <= 100 Then
                        objFormItemMaster.Items.Item("7").Specific.string = ConditionDesc + " " + SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc + " " + NetWeight + " " + NetWeightUnit + " " + BrandDesc
                    Else
                        objFormItemMaster.Items.Item("7").Specific.string = Left(ConditionDesc + " " + SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc + " " + NetWeight + " " + NetWeightUnit + " " + BrandDesc, 100)
                        objApplication.StatusBar.SetText("Max Item description 100 Digit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                End If
            ElseIf ItemGroup = "105" Or ItemGroupDescMaster = "Unpack Finish Goods" Then
                ItemGroupDesc = "UF"

                'If objFormItemMasterUDF.Items.Item("U_MISGRADE").Specific.selected.value = "" Then
                '    objApplication.StatusBar.SetText("Grade Must Fill", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    GoTo Setnothing
                'Else
                '    Grade = objFormItemMasterUDF.Items.Item("U_MISGRADE").Specific.selected.value

                '    strSql = "Select Name Descript from [@MIS_GRADE] WHERE Code = '" & Grade & "' "
                '    ObjRecSet.DoQuery(strSql)
                '    GradeDesc = ObjRecSet.Fields.Item("Descript").Value
                'End If

                ItemCodeUF = ItemGroupDesc + Brand + Species + SkinningNCut + Grade + Bagging


                strSql = "Select ISNULL(MAX(RIGHT(ItemCode,4)),0) RunNumber From OITM where ItmsGrpCod = '" & ItemGroup & "' AND ItemCode Like '" & ItemCodeUF & "%'"
                ObjRecSet.DoQuery(strSql)

                If ObjRecSet.RecordCount > 0 Then
                    intNumber = ObjRecSet.Fields.Item("RunNumber").Value
                    Select Case Len(Trim(intNumber))
                        Case 1
                            RunningNumber = "000" + CStr(Trim(intNumber + 1))
                        Case 2
                            RunningNumber = "00" + CStr(Trim(intNumber + 1))
                        Case 3
                            RunningNumber = "0" + CStr(Trim(intNumber + 1))
                            'Case 4
                            'RunningNumber = "0" + CStr(Trim(intNumber + 1))
                        Case 4
                            RunningNumber = CStr(Trim(intNumber + 1))
                    End Select
                Else
                    RunningNumber = "0001"
                End If

                intMsg = objApplication.MessageBox("Are You Sure To Generated ? Item Code =  " & ItemCodeUF + RunningNumber & " AND Description = " & SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc & "", 1, "Yes", "No", "")

                If intMsg = 1 Then
                    If Len(ItemCodeUF + RunningNumber) <= 20 Then
                        objFormItemMaster.Items.Item("5").Specific.string = ItemCodeUF + RunningNumber
                    Else
                        objFormItemMaster.Items.Item("5").Specific.string = Left(ItemCodeUF + RunningNumber, 20)
                        objApplication.StatusBar.SetText("Max Item Code 20 Digit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                    If Len(SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc) <= 100 Then
                        objFormItemMaster.Items.Item("7").Specific.string = SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc
                    Else
                        objFormItemMaster.Items.Item("7").Specific.string = Left(SpeciesDesc + " " + SkinningNCutDesc + " " + BaggingDesc + " " + TreatmentGlazDesc + " " + Sizing + " " + GradeDesc, 100)
                        objApplication.StatusBar.SetText("Max Item description 100 Digit", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    End If

                End If
            End If

            'Else
            '    objApplication.StatusBar.SetText("Input Manual Generated Code ~11.0022~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If

        Else
            objApplication.StatusBar.SetText("Generated Code Only Finish Good Or Unpack Finish Good ~11.0023~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjItemGroup)

setnothing:
        ObjItemGroup = Nothing


    End Sub

    Private Sub subLostFocusGR(ByVal TransactionType As Integer)
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        objMatrix = objFormGoodReceipt.Items.Item("13").Specific
        objColumns = objMatrix.Columns

        If TransactionType = 1 Then

            objFormGoodReceipt.Freeze(True)
            'objColumns.Item("U_MISBOXNO").Visible = True
            objColumns.Item("U_MISFISHQ").Visible = True
            objColumns.Item("U_MISINFO").Visible = False
            'objColumns.Item("U_MISFRESQ").Visible = False
            'objColumns.Item("U_MISMORTQ").Visible = False
            'objColumns.Item("U_MISUNDSQ").Visible = False
            'objColumns.Item("U_MISDEFOQ").Visible = False
            'objColumns.Item("U_MISGISQK").Visible = False
            'objColumns.Item("U_MISGISQP").Visible = False
            'objColumns.Item("U_MISVARQK").Visible = False
            'objColumns.Item("U_MISVARQP").Visible = False
            objColumns.Item("U_MISYIELD").Visible = False
            objColumns.Item("U_MISNFPRO").Visible = False
            objColumns.Item("10").Visible = True
            objColumns.Item("14").Visible = True
            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = True
            'objColumns.Item("U_MISSEALC").Visible = False

            objFormGoodReceipt.Items.Item("btnCopyGI").Visible = False

            objFormGoodReceipt.Freeze(False)
            objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000001").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISTRXNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000055").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000010").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISARRTM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000012").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000011").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000013").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000004").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000005").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000006").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000007").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000008").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000054").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000059").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000057").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000058").Visible = False

        ElseIf TransactionType = 2 Then
            objFormGoodReceipt.Freeze(True)
            objFormGoodReceipt.Items.Item("btnCopyGI").Visible = True
            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True
            objColumns.Item("U_MISINFO").Visible = False
            'objColumns.Item("U_MISFRESQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = True
            'objColumns.Item("U_MISUNDSQ").Visible = False
            'objColumns.Item("U_MISDEFOQ").Visible = False
            'objColumns.Item("U_MISGISQK").Visible = False
            'objColumns.Item("U_MISGISQP").Visible = False
            'objColumns.Item("U_MISVARQK").Visible = False
            'objColumns.Item("U_MISVARQP").Visible = False
            objColumns.Item("U_MISYIELD").Visible = False
            objColumns.Item("U_MISNFPRO").Visible = False
            objColumns.Item("10").Visible = True
            objColumns.Item("14").Visible = True
            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = False
            'objColumns.Item("U_MISNOSEG").Visible = True
            'objColumns.Item("U_MISSEALC").Visible = True

            objFormGoodReceipt.Freeze(False)
            objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000001").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISTRXNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000055").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000010").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISARRTM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000012").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000011").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000013").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000004").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000005").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000006").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000007").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000008").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000054").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000059").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000057").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000058").Visible = False

        ElseIf TransactionType = 3 Then
            objFormGoodReceipt.Freeze(True)
            objFormGoodReceipt.Items.Item("btnCopyGI").Visible = True

            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True
            objColumns.Item("U_MISINFO").Visible = True
            'objColumns.Item("U_MISFRESQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = True
            'objColumns.Item("U_MISUNDSQ").Visible = True
            'objColumns.Item("U_MISDEFOQ").Visible = True
            'objColumns.Item("U_MISGISQK").Visible = True
            'objColumns.Item("U_MISGISQP").Visible = True
            'objColumns.Item("U_MISVARQK").Visible = True
            'objColumns.Item("U_MISVARQP").Visible = True
            objColumns.Item("U_MISYIELD").Visible = False
            objColumns.Item("U_MISNFPRO").Visible = False
            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISSEALC").Visible = True

            objFormGoodReceipt.Freeze(False)
            objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000001").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISTRXNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000055").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000010").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISARRTM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000012").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000011").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000013").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000004").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000005").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000006").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000007").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000008").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000054").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000059").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000057").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000058").Visible = True

        ElseIf TransactionType = 4 Then
            objFormGoodReceipt.Freeze(True)
            objFormGoodReceipt.Items.Item("btnCopyGI").Visible = True

            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = False
            objColumns.Item("U_MISINFO").Visible = False
            'objColumns.Item("U_MISFRESQ").Visible = False
            'objColumns.Item("U_MISMORTQ").Visible = False
            'objColumns.Item("U_MISUNDSQ").Visible = False
            'objColumns.Item("U_MISDEFOQ").Visible = False
            'objColumns.Item("U_MISGISQK").Visible = False
            'objColumns.Item("U_MISGISQP").Visible = False
            'objColumns.Item("U_MISVARQK").Visible = False
            'objColumns.Item("U_MISVARQP").Visible = False
            objColumns.Item("U_MISYIELD").Visible = True
            objColumns.Item("U_MISNFPRO").Visible = True
            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = False
            'objColumns.Item("U_MISSEALC").Visible = False
            'objColumns.Item("U_MISNOSEG").Visible = False

            objFormGoodReceipt.Freeze(False)
            objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000001").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISTRXNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000055").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000010").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISARRTM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000012").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000011").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000013").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000004").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000005").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000006").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000007").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000008").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000054").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000059").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000057").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000058").Visible = False

        Else
            objFormGoodReceipt.Freeze(True)
            objFormGoodReceipt.Items.Item("btnCopyGI").Visible = False

            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = False
            objColumns.Item("U_MISINFO").Visible = False
            'objColumns.Item("U_MISFRESQ").Visible = False
            'objColumns.Item("U_MISMORTQ").Visible = False
            'objColumns.Item("U_MISUNDSQ").Visible = False
            'objColumns.Item("U_MISDEFOQ").Visible = False
            'objColumns.Item("U_MISGISQK").Visible = False
            'objColumns.Item("U_MISGISQP").Visible = False
            'objColumns.Item("U_MISVARQK").Visible = False
            'objColumns.Item("U_MISVARQP").Visible = False
            objColumns.Item("U_MISYIELD").Visible = False
            objColumns.Item("U_MISNFPRO").Visible = False
            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISSEALC").Visible = False
            objFormGoodReceipt.Freeze(False)
            objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000001").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISTRXNM").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000055").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodReceiptUDF.Items.Item("1000010").Visible = True
            objFormGoodReceiptUDF.Items.Item("U_MISARRTM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000012").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000011").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISLASTB").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000013").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000004").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000005").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000006").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000007").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000008").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISSHIFT").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000054").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000059").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000057").Visible = False
            objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Visible = False
            objFormGoodReceiptUDF.Items.Item("1000058").Visible = False
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)

    End Sub

    Private Sub subLostFocus(ByVal TransactionType As Integer)
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        objMatrix = objFormGoodIssue.Items.Item("13").Specific
        objColumns = objMatrix.Columns

        If TransactionType = 1 Then
            objFormGoodIssue.Freeze(True)
            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True

            'objColumns.Item("U_MISINFO").Cells.Item(1).Specific.select("2", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'If objColumns.Item("U_MISINFO").Cells.Item(1).Specific.selected.value = "1" Then
            '    objColumns.Item("U_MISINFO").Cells.Item(1).Specific.selected.value = "2"
            'End If

            'Dim kolom As SAPbouiCOM.Column
            'Dim kombo As SAPbouiCOM.ComboBoxColumn
            'Dim Nilai As SAPbouiCOM.ValidValue

            'kolom = objColumns.Item("U_MISINFO").Cells.Item(1).Specific
            'kombo = objColumns.Item("U_MISINFO").Cells.Item(1).Specific
            'kombo.ValidValues = Nilai




            '           Dim Tes As SAPbouiCOM.DBDataSource
            '          Tes = objFormGoodIssue.DataSources.DBDataSources.Item("IGE1")
            '         Tes.GetValue("U_MISINFO", 1)



            'objColumns.Item("U_MISINFO").Description = 2

            'objColumns.Item("U_MISMORTQ").Visible = True


            objFormGoodIssue.Items.Item("btnGen").Visible = True
            objFormGoodIssue.Items.Item("btnGen").Enabled = True
            'objFormGoodIssue.Items.Item("txtGen").Visible = True
            'objFormGoodIssue.Items.Item("txtGen").Enabled = True
            objFormGoodIssue.Items.Item("btnCopyGR").Visible = False
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = False
            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = False
            objFormGoodIssueUDF.Items.Item("1000003").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000004").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodIssueUDF.Items.Item("1000005").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000006").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodIssueUDF.Items.Item("1000007").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000008").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = False
            objFormGoodIssueUDF.Items.Item("1000009").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = False
            objFormGoodIssueUDF.Items.Item("1000010").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000011").Visible = False

        ElseIf TransactionType = 2 Then
            objFormGoodIssue.Freeze(True)

            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = False

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = False
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = False
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False

            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = False
            objFormGoodIssueUDF.Items.Item("1000003").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000004").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodIssueUDF.Items.Item("1000005").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000006").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodIssueUDF.Items.Item("1000007").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000008").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = False
            objFormGoodIssueUDF.Items.Item("1000009").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = False
            objFormGoodIssueUDF.Items.Item("1000010").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000011").Visible = False


        ElseIf TransactionType = 3 Then
            objFormGoodIssue.Freeze(True)

            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = False
            'objColumns.Item("U_MISNOSEG").Visible = True
            'objColumns.Item("U_MISBOXNO").Visible = True
            objColumns.Item("U_MISFISHQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = False

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = False
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = False
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False
            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = True
            objFormGoodIssueUDF.Items.Item("1000003").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = True
            objFormGoodIssueUDF.Items.Item("1000004").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = True
            objFormGoodIssueUDF.Items.Item("1000005").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = True
            objFormGoodIssueUDF.Items.Item("1000006").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = True
            objFormGoodIssueUDF.Items.Item("1000007").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = True
            objFormGoodIssueUDF.Items.Item("1000008").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = False
            objFormGoodIssueUDF.Items.Item("1000009").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = False
            objFormGoodIssueUDF.Items.Item("1000010").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = True
            objFormGoodIssueUDF.Items.Item("1000011").Visible = True

        ElseIf TransactionType = 4 Then
            objFormGoodIssue.Freeze(True)

            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = False

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = False
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = False
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False

            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = False
            objFormGoodIssueUDF.Items.Item("1000003").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000004").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodIssueUDF.Items.Item("1000005").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000006").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodIssueUDF.Items.Item("1000007").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000008").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = True
            objFormGoodIssueUDF.Items.Item("1000009").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = False
            objFormGoodIssueUDF.Items.Item("1000010").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000011").Visible = False

        ElseIf TransactionType = 5 Then
            objFormGoodIssue.Freeze(True)
            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = True
            'objColumns.Item("U_MISBOXNO").Visible = True
            objColumns.Item("U_MISFISHQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = False

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = True
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = True
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False
            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = True
            objFormGoodIssueUDF.Items.Item("1000003").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = True
            objFormGoodIssueUDF.Items.Item("1000004").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = True
            objFormGoodIssueUDF.Items.Item("1000005").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = True
            objFormGoodIssueUDF.Items.Item("1000006").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = True
            objFormGoodIssueUDF.Items.Item("1000007").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = True
            objFormGoodIssueUDF.Items.Item("1000008").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = False
            objFormGoodIssueUDF.Items.Item("1000009").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodIssueUDF.Items.Item("1000010").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = True
            objFormGoodIssueUDF.Items.Item("1000011").Visible = True

        ElseIf TransactionType = 6 Then
            objFormGoodIssue.Freeze(True)
            objColumns.Item("U_MISNETID").Visible = False
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISBOXNO").Visible = True
            objColumns.Item("U_MISFISHQ").Visible = True
            'objColumns.Item("U_MISMORTQ").Visible = True

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = True
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = True
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False

            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = False
            objFormGoodIssueUDF.Items.Item("1000003").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000004").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodIssueUDF.Items.Item("1000005").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000006").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodIssueUDF.Items.Item("1000007").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000008").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = False
            objFormGoodIssueUDF.Items.Item("1000009").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodIssueUDF.Items.Item("1000010").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000011").Visible = False



        Else
            objFormGoodIssue.Freeze(True)
            objColumns.Item("U_MISNETID").Visible = True
            objColumns.Item("U_MISPROID").Visible = True
            'objColumns.Item("U_MISNOSEG").Visible = False
            'objColumns.Item("U_MISBOXNO").Visible = False
            objColumns.Item("U_MISFISHQ").Visible = True

            objFormGoodIssue.Items.Item("btnCopyGR").Visible = False
            objFormGoodIssue.Items.Item("btnCopyGR").Enabled = False
            objFormGoodIssue.Items.Item("btnGen").Visible = False
            'objFormGoodIssue.Items.Item("txtGen").Visible = False
            objFormGoodIssue.Freeze(False)

            objFormGoodIssueUDF.Items.Item("U_MISDESTW").Visible = True
            objFormGoodIssueUDF.Items.Item("1000003").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000004").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISASDRV").Visible = False
            objFormGoodIssueUDF.Items.Item("1000005").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISLICNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000006").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISSPVID").Visible = False
            objFormGoodIssueUDF.Items.Item("1000007").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISRITNO").Visible = False
            objFormGoodIssueUDF.Items.Item("1000008").Visible = False
            objFormGoodIssueUDF.Items.Item("U_MISREASC").Visible = True
            objFormGoodIssueUDF.Items.Item("1000009").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISREFFD").Visible = True
            objFormGoodIssueUDF.Items.Item("1000010").Visible = True
            objFormGoodIssueUDF.Items.Item("U_MISDELTM").Visible = False
            objFormGoodIssueUDF.Items.Item("1000011").Visible = False

        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)

    End Sub

    Private Sub subFPFindData(ByVal Project As String, ByVal pForm As SAPbouiCOM.Form, ByRef pBubbleEvent As Boolean)
        Dim objRecSet As SAPbobsCOM.Recordset
        Dim StrSql As String
        Dim strFind As String



        StrSql = "Select Top 1 T0.* " & _
                "From [@MIS_PRJMSTR] T0"

        If (Trim(pForm.Items.Item("MISNETID").Specific.String) <> "" And Trim(pForm.Items.Item("MISSIGND").Specific.String) <> "") Then
            strFind = strFind & " T0.U_MISNETID = '" & pForm.Items.Item("MISNETID").Specific.String & "' And T0.U_MISSIGND = '" & CDate(ClsGlobal.fctFormatDateSave(oCompany, pForm.Items.Item("MISSIGND").Specific.String, 1)) & "' And "
        End If

        If strFind <> "" Then
            strFind = " Where " & strFind
        End If

        If Right(strFind, 4) = "And " Then
            strFind = Left(strFind, Len(strFind) - 4)
        End If

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If strFind = "" Then
            pBubbleEvent = False
            objApplication.StatusBar.SetText("No matching records found 'Net Id' ~10.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo Setnothing
        Else
            objRecSet.DoQuery(StrSql & strFind & " Order by T0.U_MISPROID")
        End If

        If objRecSet.RecordCount > 0 Then
            SubProjectMasterDisplayData(Project, objRecSet, pForm, UCase(pForm.Items.Item("btnOK").Specific.Caption))
        Else
            pBubbleEvent = False
            objApplication.StatusBar.SetText("No matching records found 'Net Id' ~10.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


Setnothing:
        objRecSet = Nothing

    End Sub

    Private Function fctValidateMortal(ByVal pForm As SAPbouiCOM.Form, ByRef pBubbleEvent As Boolean) As Boolean
        'Dim strSql As String
        Dim BubbleEvent As Boolean
        Dim ObjRecSet As SAPbobsCOM.Recordset = Nothing
        Dim ProjectmortQty As Integer
        Dim mortalQty As Integer
        Dim mortalDate As String
        Dim Project As String
        Dim intMsg As Integer

        fctValidateMortal = False
        pBubbleEvent = False

        'ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        'strSql = "SELECT U_MISPROID FROM [@MIS_PRJMSTR] Where U_MISPROID = '" & pForm.Items.Item("NetCd").Specific.string & " '"
        'ObjRecSet.DoQuery(strSql)

        If pForm.Items.Item("PrjCd").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill Project Code ~5.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("DtDead").Specific.value = "" Then
            objApplication.StatusBar.SetText("U Must Fill Mortality Date ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("MortQty").Specific.value = "" Then
            objApplication.StatusBar.SetText("U Must Fill Mortality Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

        fctValidateMortal = True
        pBubbleEvent = True

        Project = pForm.Items.Item("PrjCd").Specific.string
        ProjectmortQty = objFormProjectMaster.Items.Item("MISNFDIE").Specific.value
        mortalQty = pForm.Items.Item("MortQty").Specific.value
        mortalDate = pForm.Items.Item("DtDead").Specific.value


        intMsg = objApplication.MessageBox("Are You Sure To Add Mortality Fish " & ProjectmortQty & " + " & mortalQty & " = " & ProjectmortQty + mortalQty & " ", 1, "OK", "CANCEL")

        If intMsg = 1 Then
            If Not fctTestingSaveMortal(oFormLookUpMortal, "ADD", BubbleEvent, Project, ProjectmortQty + mortalQty, mortalDate) Then GoTo Setnothing
            pForm.Close()
            objFormProjectMaster.Items.Item("MISNFDIE").Enabled = True
            objFormProjectMaster.Items.Item("MISNFDIE").Specific.value = ProjectmortQty + mortalQty
            objFormProjectMaster.Items.Item("MISNETID").Click()
            objFormProjectMaster.Items.Item("MISNFDIE").Enabled = False
            If Not fctTestingSave(objFormProjectMaster, "UPDATE", BubbleEvent) Then GoTo Setnothing
        Else
            fctValidateMortal = True
            pBubbleEvent = True
        End If
setnothing:
        '        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)

    End Function

    Private Function fctValidateBatchFG(ByVal pForm As SAPbouiCOM.Form, ByRef pBubbleEvent As Boolean) As Boolean

        fctValidateBatchFG = False
        pBubbleEvent = False

        If pForm.Items.Item("Region").Specific.string = "" Then
            objApplication.StatusBar.SetText("Region Must Be Fill ~4.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("DocDate").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill Date ~4.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

        fctValidateBatchFG = True
        pBubbleEvent = True

setnothing:
    End Function

    Private Function fctValidateBatchRM(ByVal pForm As SAPbouiCOM.Form, ByRef pBubbleEvent As Boolean) As Boolean
        Dim strSql As String
        Dim ObjRecSet As SAPbobsCOM.Recordset = Nothing

        fctValidateBatchRM = False
        pBubbleEvent = False


        ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strSql = "SELECT U_MISNETCD FROM [@MIS_NETMS] Where U_MISNETCD = '" & pForm.Items.Item("NetCd").Specific.string & " '"
        ObjRecSet.DoQuery(strSql)

        If pForm.Items.Item("NetCd").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill NET Code ~5.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf ObjRecSet.RecordCount = 0 Then
            objApplication.StatusBar.SetText("Net Id Not Available in Net Master, Please Input Net Master ~5.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("DocDate").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill Date ~5.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf Len(pForm.Items.Item("NetCd").Specific.string) <> 8 Then
            objApplication.StatusBar.SetText("NET Code Length Must 8 Digit ~5.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("RitNo").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill Rit No ~5.0005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf Len(pForm.Items.Item("RitNo").Specific.string) <> 2 Then
            objApplication.StatusBar.SetText("Rit No Length Must 2 Digit ~5.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf pForm.Items.Item("BoxNo").Specific.string = "" Then
            objApplication.StatusBar.SetText("U Must Fill Box No ~5.0007~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        ElseIf Len(pForm.Items.Item("BoxNo").Specific.string) <> 2 Then
            objApplication.StatusBar.SetText("Box No Length Must 2 Digit ~5.0008~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

        fctValidateBatchRM = True
        pBubbleEvent = True
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)

setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)

    End Function

    Private Function fctValidate(ByVal pForm As SAPbouiCOM.Form, ByVal pMode As String, ByRef pBubbleEvent As Boolean) As Boolean
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim strsql As String
        Dim ProjectCd As String
        Dim NetCode As String
        Dim Species As String
        Dim FCR As Double
        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        If pMode = "ADD" Then

            If objFormProjectMaster.Items.Item("MISNETID").Specific.string = "" Then
                objApplication.StatusBar.SetText("Net Code Must Be Fill. ~2.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            ElseIf objFormProjectMaster.Items.Item("MISNETID").Specific.string <> "" Then
                strsql = "SELECT U_MISNETID FROM [@MIS_PRJMSTR] WHERE U_MISNETID = '" & objFormProjectMaster.Items.Item("MISNETID").Specific.string & "' AND (U_MISNETST = 'O' OR U_MISNETST = 'H') "
                objRecSet.DoQuery(strsql)

                If objRecSet.RecordCount <> 0 Then
                    objApplication.StatusBar.SetText("U Must First Close Project In Same Net Master", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    GoTo Setnothing
                End If
            End If

            If objFormProjectMaster.Items.Item("MISSIGND").Specific.string = "" Then
                objApplication.StatusBar.SetText("Stocking Date Must Be Fill. ~2.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISNETID").Specific.string <> Left(objFormProjectMaster.Items.Item("MISNETID").Specific.string, 8) Then
                objApplication.StatusBar.SetText("Project Code Must Same With Net Code, Please Lost Focus Stocking Date. ~2.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISPROID").Specific.string = "" Then
                objApplication.StatusBar.SetText("Project Id Must Be Fill. ~2.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISSCIES").Specific.string = "" Then
                objApplication.StatusBar.SetText("Species Must Be Fill. ~2.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISHARVP").Specific.string = "" _
                Or objFormProjectMaster.Items.Item("MISHARVP").Specific.string <= 0 Then
                objApplication.StatusBar.SetText("Day Of Culture In Days > 0. ~2.0005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISESTSF").Specific.value = 0 Then
                objApplication.StatusBar.SetText("Estimated No Of Fish Must Be Fill. ~2.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISESTHQ").Specific.value = "" _
            Or objFormProjectMaster.Items.Item("MISESTHQ").Specific.value = 0 Then
                objApplication.StatusBar.SetText("Estimated Harvesting Qty Must Be Fill. ~2.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If objFormProjectMaster.Items.Item("MISESTHD").Specific.string = "" Then
                objApplication.StatusBar.SetText("Estimated Harvesting Date Must Be Fill, Please Lost Focus Stock Date. ~2.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            Dim EstimasiFish As Integer
            Dim EstimasiHarvest As Integer

            EstimasiFish = objFormProjectMaster.Items.Item("MISESTSF").Specific.value * _
                objFormProjectMaster.Items.Item("MISESTLF").Specific.value / 100

            EstimasiHarvest = objFormProjectMaster.Items.Item("MISESTHQ").Specific.value

            If EstimasiFish <> EstimasiHarvest Then
                objApplication.StatusBar.SetText("Estimated No Of Fish Must Lost Focus. ~2.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If



            ProjectCd = objFormProjectMaster.Items.Item("MISPROID").Specific.String
            NetCode = objFormProjectMaster.Items.Item("MISNETID").Specific.String
            Species = objFormProjectMaster.Items.Item("MISSCIES").Specific.string

            If ProjectCd <> "" Then
                strsql = "Select U_MISPROID From [@MIS_PRJMSTR] Where U_MISPROID = '" & ProjectCd & "' AND U_MISNETST <> 'D' "
                objRecSet.DoQuery(strsql)

                If objRecSet.RecordCount > 0 Then
                    objApplication.StatusBar.SetText("Project Id Already exists. ~2.0007~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    GoTo Setnothing
                End If
            End If

            If objFormProjectMaster.Items.Item("MISFCR").Specific.value = 0 Then
                objApplication.StatusBar.SetText("Feed Consumtion Rate Must Fill, Please Check Rate Master. ~2.0007~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            If NetCode <> "" Then
                strsql = "Select U_MISNETCD From [@MIS_NETMS] Where U_MISRECST = 'O' AND U_MISNETCD = '" & NetCode & "' "
                objRecSet.DoQuery(strsql)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Net Code Must be Active. ~2.0008~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    GoTo Setnothing
                End If
            End If

            If Species <> "" Then
                strsql = "Select Code From [@MIS_SPEC] Where U_MISRECST = 'A' AND Code = '" & Species & "' "
                objRecSet.DoQuery(strsql)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Species Must be Active. ~2.0009~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    GoTo Setnothing
                End If
            End If


        ElseIf pMode = "UPDATE" Then
            If objFormProjectMaster.Items.Item("MISNETID").Specific.string = "" Then
                objApplication.StatusBar.SetText("Net Code Must Be Fill. ~3.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            ElseIf objFormProjectMaster.Items.Item("MISSIGND").Specific.string = "" Then
                objApplication.StatusBar.SetText("Stocking Date Must Be Fill. ~3.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            ElseIf objFormProjectMaster.Items.Item("MISPROID").Specific.string = "" Then
                objApplication.StatusBar.SetText("Project Id Must Be Fill. ~3.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            ElseIf objFormProjectMaster.Items.Item("MISHARVP").Specific.string = "" _
                Or objFormProjectMaster.Items.Item("MISHARVP").Specific.string <= 0 Then
                objApplication.StatusBar.SetText("Day Of Culture In Days > 0. ~3.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GoTo Setnothing
            End If

            ProjectCd = objFormProjectMaster.Items.Item("MISPROID").Specific.String
            NetCode = objFormProjectMaster.Items.Item("MISNETID").Specific.String

            'If ProjectCd <> "" Then
            '    strsql = "Select U_MISPROID From [@MIS_PRJMSTR] Where U_MISPROID = '" & ProjectCd & "' "
            '    objRecSet.DoQuery(strsql)

            '    If objRecSet.RecordCount > 0 Then
            '        objApplication.StatusBar.SetText("Project Id Already exists. ~30005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        GoTo Setnothing
            '    End If
            'End If

            If NetCode <> "" Then
                strsql = "Select U_MISNETCD From [@MIS_NETMS] Where U_MISRECST = 'O' AND U_MISNETCD = '" & NetCode & "' "
                objRecSet.DoQuery(strsql)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Net Code Must be Active. ~3.0005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    GoTo Setnothing
                End If
            End If

        End If

        fctValidate = True
        pBubbleEvent = True


Setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        'objRecSet = Nothing
    End Function

    Private Function fctTestingSaveMortal(ByVal pForm As SAPbouiCOM.Form, ByVal pMode As String, ByRef pBubbleEvent As Boolean, ByVal Project As String, ByVal MortalQty As Double, ByVal mortaldate As String) As Boolean
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objRecSetDetail As SAPbobsCOM.Recordset = Nothing
        Dim intRow As Integer
        Dim intDocEntry As Integer
        Dim intDocnum As Integer
        Dim Period As Integer
        Dim Seri As Integer
        Dim UserSign As String
        Dim CreateDate As String
        Dim CreateTime As Integer

        Dim DocEntry As String
        Dim Kode As Integer
        Dim StrKode As String
        Dim intLine As Integer
        Dim StrSql As String
        Dim StrSQLDetail As String
        Dim Reason As String

        Dim ProjectCd As String
        Dim MortalityDate As Date
        Dim MortalityQty As Integer

        ProjectCd = oFormLookUpMortal.Items.Item("PrjCd").Specific.String
        MortalityQty = oFormLookUpMortal.Items.Item("MortQty").Specific.String
        Reason = oFormLookUpMortal.Items.Item("Reason").Specific.String
        If oFormLookUpMortal.Items.Item("DtDead").Specific.String = "" Then
            MortalityDate = "12:00:00 AM"
        Else
            MortalityDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, oFormLookUpMortal.Items.Item("DtDead").Specific.String, 1))
        End If

        fctTestingSaveMortal = False
        pBubbleEvent = False


        On Error GoTo ErrorHandler

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Call oCompany.StartTransaction()

        StrSql = "SELECT (CASE " & _
        "WHEN (SELECT MAX(DocEntry) FROM [@MIS_FISHMORT]) IS NULL THEN 1 " & _
        "ELSE (SELECT MAX(DocEntry) + 1 FROM [@MIS_FISHMORT]) END " & _
        ") Code, (CASE  " & _
        "WHEN (SELECT NEXTNUMBER FROM NNM1 Where ObjectCode = 'MISMORTQ' And Series =  " & _
        "(SELECT DfltSeries FROM ONNM WHERE ObjectCode = 'MISMORTQ')) IS NULL THEN 1 " & _
        "ELSE (SELECT NEXTNUMBER FROM NNM1 Where ObjectCode = 'MISMORTQ' And Series = " & _
        "(SELECT DfltSeries FROM ONNM WHERE ObjectCode = 'MISMORTQ')) END) DocNum, " & _
        "(CASE WHEN(SELECT AbsEntry FROM OFPR WHERE CONVERT(VARCHAR(8),getdate(),112) BETWEEN F_RefDate AND T_RefDate)IS NULL THEN " & _
        "(SELECT TOP 1 AbsEntry FROM OFPR WHERE GETDATE() BETWEEN F_DueDate AND T_DueDate ORDER BY AbsEntry Desc) " & _
        "ELSE(SELECT AbsEntry FROM OFPR WHERE CONVERT(VARCHAR(8),getdate(),112) BETWEEN F_RefDate AND T_RefDate) " & _
        "END) Period, " & _
        "(CASE WHEN(SELECT Series FROM NNM1 WHERE ObjectCode = 'MISMORTQ') IS NULL THEN " & _
        "(SELECT TOP 1 Series FROM NNM1 ORDER BY NNM1.Series DESC) " & _
        "ELSE (SELECT Series FROM NNM1 WHERE ObjectCode = 'MISMORTQ') END) Series, " & _
        "(CONVERT(VARCHAR, GETDATE(),112)) CreateDate, (DATEPART(HH, GETDATE()) * 100 + DATEPART(MI, GETDATE())) CreateTime"

        objRecSet.DoQuery(StrSql)

        If objRecSet.RecordCount > 0 Then
            intDocEntry = objRecSet.Fields.Item("Code").Value
            intDocnum = objRecSet.Fields.Item("DocNum").Value
            Period = objRecSet.Fields.Item("Period").Value
            Seri = objRecSet.Fields.Item("Series").Value
            CreateDate = objRecSet.Fields.Item("CreateDate").Value
            CreateTime = objRecSet.Fields.Item("CreateTime").Value
            If pMode = "ADD" Then
                UserSign = oCompany.UserSignature

                StrSQLDetail = "Insert Into [@MIS_FISHMORT]" & _
                                "([DocEntry], [DocNum],[Period],[Series],[Object], [UserSign],[CreateDate],[CreateTime],[DataSource], " & _
                                " [U_MISPROID], [U_MISMORQT], [U_MISMORDT], [U_MISREASC]) " & _
                                "values(" & intDocEntry & ", " & intDocnum & "," & Period & "," & Seri & ", 'PrjMstr','" & UserSign & "','" & CreateDate & "','" & CreateTime & "','I' , '" & ProjectCd & "','" & MortalityQty & "'," & IIf(MortalityDate = "12:00:00 AM", "NULL", "'" & MortalityDate & "'") & ",'" & Reason & "' )"

                objRecSet.DoQuery(StrSQLDetail)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Insert Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                Else
                    objApplication.StatusBar.SetText("Contact IT Support ~1.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                End If


                StrSQLDetail = " UPDATE NNM1 " & _
                                "SET NextNumber = " & intDocnum + 1 & " " & _
                                "WHERE ObjectCode = 'MISMORTQ' AND Series = " & Seri & ""
                objRecSet.DoQuery(StrSQLDetail)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Update Successfull", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                Else
                    objApplication.StatusBar.SetText("Contact IT Support ~1.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                End If

            End If

        Else
            objApplication.StatusBar.SetText("Wrong Format Contact IT Support ~1.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo Setnothing
        End If

        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

        If pMode = "ADD" Then
            objApplication.StatusBar.SetText("Operation completed successfully. (Doc.Entry No.: " & Trim(CStr(intDocEntry)) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objApplication.StatusBar.SetText("Operation Update completed successfully.(Doc.Entry No.: " & Trim(CStr(DocEntry)) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        fctTestingSaveMortal = True
        pBubbleEvent = True

        GoTo Setnothing


ErrorHandler:
        If Err.Number <> 0 Then
            If oCompany.InTransaction Then
                Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            objApplication.StatusBar.SetText("Fail saving data.(" & Trim(oCompany.GetLastErrorDescription) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

Setnothing:
        If oCompany.InTransaction Then
            Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        objRecSet = Nothing

    End Function

    Private Function fctTestingSave(ByVal pForm As SAPbouiCOM.Form, ByVal pMode As String, ByRef pBubbleEvent As Boolean) As Boolean
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objRecSetDetail As SAPbobsCOM.Recordset = Nothing
        Dim intRow As Integer
        Dim intDocEntry As Integer
        Dim intDocnum As Integer
        Dim Period As Integer
        Dim Seri As Integer
        Dim UserSign As String
        Dim CreateDate As String
        Dim CreateTime As Integer

        Dim DocEntry As String
        Dim Kode As Integer
        Dim StrKode As String
        Dim intLine As Integer
        Dim StrSql As String
        Dim StrSQLDetail As String

        Dim ProjectCd As String
        Dim NetCode As String
        Dim StockingDate As Date
        Dim Species As String
        Dim EstimatedNoFish As Integer
        Dim DayofCulture As Integer
        Dim AgeTransferFish As Integer
        Dim EstimatedHarvestDate As Date
        Dim SurvivalRate As Double
        Dim EstimatedHarvestQty As Integer
        Dim CumulativeMortality As Integer
        Dim NetPurposes As String
        Dim FingerlingBatchCode As String
        Dim ProjectStockingRemarks As String
        Dim ActualHarvestDate As Date
        Dim ProjectHarvestRemarks As String
        Dim initFishQtyKg As Double
        Dim FeedConsumptionKg As Double
        Dim FeedConsumptionRatio As Double
        Dim FeedConsumptionEstimated As Double
        Dim TotalEstimatedFishInKg As Double
        Dim initFishCost As Double
        Dim FeedConsumptionCost As Double
        Dim TotalCostEstimated As Double
        Dim TotalGRCostActual As Double
        Dim TotalGRQtyActual As Double
        Dim ProjectCalculationFlag As String
        Dim NetStatus As String
        Dim GeneticCode As String
        Dim ActualHarvestQty As Double
        Dim Klasifikasi As String

        ProjectCd = objFormProjectMaster.Items.Item("MISPROID").Specific.String
        NetCode = objFormProjectMaster.Items.Item("MISNETID").Specific.String
        If objFormProjectMaster.Items.Item("MISSIGND").Specific.String = "" Then
            StockingDate = "12:00:00 AM"
        Else
            StockingDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISSIGND").Specific.String, 1)) 'CDate(objFormProjectMaster.Items.Item("MISSIGND").Specific.string) '
        End If
        Species = objFormProjectMaster.Items.Item("MISSCIES").Specific.String
        EstimatedNoFish = objFormProjectMaster.Items.Item("MISESTSF").Specific.value
        DayofCulture = objFormProjectMaster.Items.Item("MISHARVP").Specific.value
        If objFormProjectMaster.Items.Item("MISAGETR").Specific.String = "" Then
            AgeTransferFish = 0
        Else
            AgeTransferFish = objFormProjectMaster.Items.Item("MISAGETR").Specific.value
        End If

        If objFormProjectMaster.Items.Item("MISESTHD").Specific.String = "" Then
            EstimatedHarvestDate = "12:00:00 AM"
        Else
            EstimatedHarvestDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISESTHD").Specific.String, 2))
        End If
        '        GeneticCode = objFormProjectMaster.Items.Item("MISGENET").Specific.String
        SurvivalRate = objFormProjectMaster.Items.Item("MISESTLF").Specific.value

        If objFormProjectMaster.Items.Item("MISESTHQ").Specific.String = "" Then
            EstimatedHarvestQty = 0
        Else
            EstimatedHarvestQty = objFormProjectMaster.Items.Item("MISESTHQ").Specific.value
        End If

        CumulativeMortality = objFormProjectMaster.Items.Item("MISNFDIE").Specific.value
        NetPurposes = objFormProjectMaster.Items.Item("MISNETPUCD").Specific.String
        If objFormProjectMaster.Items.Item("MISGENCD").Specific.String = "" Then
            FingerlingBatchCode = 0
        Else
            FingerlingBatchCode = objFormProjectMaster.Items.Item("MISGENCD").Specific.String
        End If

        ProjectStockingRemarks = objFormProjectMaster.Items.Item("MISPROSR").Specific.String
        If objFormProjectMaster.Items.Item("MISHARVD").Specific.String = "" Then
            ActualHarvestDate = "12:00:00 AM"
        Else
            ActualHarvestDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, objFormProjectMaster.Items.Item("MISHARVD").Specific.String, 3))
        End If

        If objFormProjectMaster.Items.Item("MISHARVQ").Specific.String = "" Then
            ActualHarvestQty = 0
        Else
            ActualHarvestQty = objFormProjectMaster.Items.Item("MISHARVQ").Specific.String
        End If
        ProjectHarvestRemarks = objFormProjectMaster.Items.Item("MISPROHR").Specific.String
        initFishQtyKg = objFormProjectMaster.Items.Item("MISINIFQ").Specific.value
        Klasifikasi = objFormProjectMaster.Items.Item("MISHATGO").Specific.String
        FeedConsumptionKg = objFormProjectMaster.Items.Item("MISFEEDQ").Specific.value
        FeedConsumptionRatio = objFormProjectMaster.Items.Item("MISFCR").Specific.value
        FeedConsumptionEstimated = objFormProjectMaster.Items.Item("MISFCE").Specific.value
        TotalEstimatedFishInKg = objFormProjectMaster.Items.Item("MISTEFQK").Specific.value
        initFishCost = objFormProjectMaster.Items.Item("MISINIFC").Specific.value
        FeedConsumptionCost = objFormProjectMaster.Items.Item("MISFEEDC").Specific.value
        TotalCostEstimated = objFormProjectMaster.Items.Item("MISTPCST").Specific.value
        TotalGRCostActual = objFormProjectMaster.Items.Item("MISTPGRC").Specific.value
        TotalGRQtyActual = objFormProjectMaster.Items.Item("MISTPGRQ").Specific.value
        ProjectCalculationFlag = objFormProjectMaster.Items.Item("MISPROCS").Specific.String
        NetStatus = "O" 'objFormProjectMaster.Items.Item("MISNETST").Specific.String



        fctTestingSave = False
        pBubbleEvent = False


        On Error GoTo ErrorHandler

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Call oCompany.StartTransaction()

        StrSql = "SELECT (CASE " & _
        "WHEN (SELECT MAX(DocEntry) FROM [@MIS_PRJMSTR]) IS NULL THEN 1 " & _
        "ELSE (SELECT MAX(DocEntry) + 1 FROM [@MIS_PRJMSTR]) END " & _
        ") Code, (CASE  " & _
        "WHEN (SELECT NEXTNUMBER FROM NNM1 Where ObjectCode = 'PrjMstr' And Series =  " & _
        "(SELECT DfltSeries FROM ONNM WHERE ObjectCode = 'PrjMstr')) IS NULL THEN 1 " & _
        "ELSE (SELECT NEXTNUMBER FROM NNM1 Where ObjectCode = 'PrjMstr' And Series = " & _
        "(SELECT DfltSeries FROM ONNM WHERE ObjectCode = 'PrjMstr')) END) DocNum, " & _
        "(CASE WHEN(SELECT AbsEntry FROM OFPR WHERE CONVERT(VARCHAR(8),getdate(),112) BETWEEN F_RefDate AND T_RefDate)IS NULL THEN " & _
        "(SELECT TOP 1 AbsEntry FROM OFPR WHERE GETDATE() BETWEEN F_DueDate AND T_DueDate ORDER BY AbsEntry Desc) " & _
        "ELSE(SELECT AbsEntry FROM OFPR WHERE CONVERT(VARCHAR(8),getdate(),112) BETWEEN F_RefDate AND T_RefDate) " & _
        "END) Period, " & _
        "(CASE WHEN(SELECT Series FROM NNM1 WHERE ObjectCode = 'PrjMstr') IS NULL THEN " & _
        "(SELECT TOP 1 Series FROM NNM1 ORDER BY NNM1.Series DESC) " & _
        "ELSE (SELECT Series FROM NNM1 WHERE ObjectCode = 'PrjMstr') END) Series, " & _
        "(CONVERT(VARCHAR, GETDATE(),112)) CreateDate, (DATEPART(HH, GETDATE()) * 100 + DATEPART(MI, GETDATE())) CreateTime"


        objRecSet.DoQuery(StrSql)

        If objRecSet.RecordCount > 0 Then
            intDocEntry = objRecSet.Fields.Item("Code").Value
            intDocnum = objRecSet.Fields.Item("DocNum").Value
            Period = objRecSet.Fields.Item("Period").Value
            Seri = objRecSet.Fields.Item("Series").Value
            CreateDate = objRecSet.Fields.Item("CreateDate").Value
            CreateTime = objRecSet.Fields.Item("CreateTime").Value


            If pMode = "ADD" Then

                'StrSql = "Select ISNULL(MAX(code),0)+1 Code From [@MIS_PRJMSTR]"
                'objRecSet.DoQuery(StrSql)

                'If objRecSet.RecordCount > 0 Then
                '    intDocEntry = objRecSet.Fields.Item("Code").Value
                '    '    Select Case Len(Trim(intDocEntry))
                '    '        Case 1
                '    '            StrKode = "0000000" + CStr(Trim(intDocEntry))
                '    '        Case 2
                '    '            StrKode = "000000" + CStr(Trim(intDocEntry))
                '    '        Case 3
                '    '            StrKode = "00000" + CStr(Trim(intDocEntry))
                '    '        Case 4
                '    '            StrKode = "0000" + CStr(Trim(intDocEntry))
                '    '        Case 5
                '    '            StrKode = "000" + CStr(Trim(intDocEntry))
                '    '        Case 6
                '    '            StrKode = "00" + CStr(Trim(intDocEntry))
                '    '        Case 7
                '    '            StrKode = "0" + CStr(Trim(intDocEntry))
                '    '        Case 8
                '    '            StrKode = CStr(Trim(intDocEntry))
                '    '    End Select
                '    'Else
                '    '    StrKode = "00000001"
                'End If



                UserSign = oCompany.UserSignature

                StrSQLDetail = "Insert Into [@MIS_PRJMSTR]" & _
                                "([DocEntry], [DocNum],[Period],[Series],[Object], [UserSign],[CreateDate],[CreateTime],[DataSource], [U_MISPROID],[U_MISNETID], [U_MISSIGND], [U_MISSCIES], [U_MISESTSF], [U_MISHARVP], " & _
                                " [U_MISAGETR], [U_MISESTHD], [U_MISESTLF], [U_MISESTHQ], [U_MISNFDIE], [U_MISNETPU], [U_MISGENCD], " & _
                                " [U_MISPROSR], [U_MISHARVD], [U_MISPROHR], [U_MISINIFQ], [U_MISFEEDQ], [U_MISFCR],[U_MISFCE],[U_MISTFQKG],[U_MISINIFC], [U_MISFEEDC], " & _
                                " [U_MISTPCST], [U_MISTPGRC], [U_MISTPGRQ], [U_MISPROCS], [U_MISNETST], [U_MISHARVQ], [U_MISHATGRO])" & _
                                "values(" & intDocEntry & ", " & intDocnum & "," & Period & "," & Seri & ", 'PrjMstr','" & UserSign & "','" & CreateDate & "','" & CreateTime & "','I' , '" & ProjectCd & "','" & NetCode & "'," & IIf(StockingDate = "12:00:00 AM", "NULL", "'" & StockingDate & "'") & " " & _
                                ",'" & Species & "'," & EstimatedNoFish & "," & DayofCulture & "," & AgeTransferFish & "," & IIf(EstimatedHarvestDate = "12:00:00 AM", "NULL", "'" & EstimatedHarvestDate & "'") & " " & _
                                "," & SurvivalRate & "," & EstimatedHarvestQty & "," & CumulativeMortality & ",'" & NetPurposes & "','" & FingerlingBatchCode & "'" & _
                                ",'" & ProjectStockingRemarks & "'," & IIf(ActualHarvestDate = "12:00:00 AM", "NULL", "'" & ActualHarvestDate & "'") & " " & _
                                ",'" & ProjectHarvestRemarks & "', " & initFishQtyKg & "," & FeedConsumptionKg & "," & FeedConsumptionRatio & "," & FeedConsumptionEstimated & "," & TotalEstimatedFishInKg & ", " & initFishCost & ", " & FeedConsumptionCost & " " & _
                                "," & TotalCostEstimated & "," & TotalGRCostActual & ", " & TotalGRQtyActual & ",'" & ProjectCalculationFlag & "','" & NetStatus & "', " & ActualHarvestQty & ", '" & Klasifikasi & "' )"

                objRecSet.DoQuery(StrSQLDetail)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Insert Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                Else
                    objApplication.StatusBar.SetText("Contact IT Support ~1.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                End If


                StrSQLDetail = " UPDATE NNM1 " & _
                                "SET NextNumber = " & intDocnum + 1 & " " & _
                                "WHERE ObjectCode = 'PrjMstr' AND Series = " & Seri & ""
                objRecSet.DoQuery(StrSQLDetail)

                If objRecSet.RecordCount = 0 Then
                    objApplication.StatusBar.SetText("Update Successfull", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                Else
                    objApplication.StatusBar.SetText("Contact IT Support ~1.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GoTo Setnothing
                End If


                'StrSql = "SELECT PrjCode FROM OPRJ WHERE PrjCode = '" & NetCode & "'"
                'objRecSet.DoQuery(StrSql)

                'If objRecSet.RecordCount > 0 Then
                '    StrSql = "Insert Into OPRJ" & _
                '            "([PrjCode], [PrjName]) VALUES ('" & NetCode & "','" & ProjectCd & "' )"
                '    objRecSet.DoQuery(StrSql)
                'Else
                '    objApplication.StatusBar.SetText("Contact IT Support ~1.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    GoTo Setnothing
                'End If


            ElseIf pMode = "UPDATE" Then

                DocEntry = objFormProjectMaster.Items.Item("Code").Specific.string

                StrSql = "Update [@MIS_PRJMSTR] " & _
                        "Set [UserSign] = '" & UserSign & "', UpdateDate = '" & CreateDate & "', UpdateTime = '" & CreateTime & "', U_MISPROID = '" & ProjectCd & "', U_MISNETID = '" & NetCode & "', U_MISSIGND = " & IIf(StockingDate = "12:00:00 AM", "NULL", "'" & StockingDate & "'") & ", " & _
                        "U_MISSCIES = '" & Species & "', U_MISESTSF = " & EstimatedNoFish & ", " & _
                        "U_MISHARVP = " & DayofCulture & ", " & _
                        "U_MISAGETR = " & AgeTransferFish & ", " & _
                        "U_MISESTHD = " & IIf(EstimatedHarvestDate = "12:00:00 AM", "NULL", "'" & EstimatedHarvestDate & "'") & ", " & _
                        "U_MISESTLF = " & SurvivalRate & ", " & _
                        "U_MISESTHQ = " & EstimatedHarvestQty & ", U_MISNFDIE = " & CumulativeMortality & ", " & _
                        "U_MISNETPU = '" & NetPurposes & "', " & _
                        "U_MISGENCD = '" & FingerlingBatchCode & "', " & _
                        "U_MISPROSR = '" & ProjectStockingRemarks & "',U_MISHARVD = " & IIf(ActualHarvestDate = "12:00:00 AM", "NULL", "'" & ActualHarvestDate & "'") & ", " & _
                        "U_MISPROHR = '" & ProjectHarvestRemarks & "', " & _
                        "U_MISINIFQ = " & initFishCost & ", " & _
                        "U_MISFEEDQ = " & FeedConsumptionCost & ",U_MISFCR = " & FeedConsumptionCost & ",U_MISFCE = " & FeedConsumptionEstimated & ", U_MISTFQKG = " & TotalEstimatedFishInKg & "," & _
                        "U_MISINIFC = " & initFishCost & ", " & _
                        "U_MISFEEDC = " & FeedConsumptionCost & ", " & _
                        "U_MISTPCST = " & TotalCostEstimated & ", " & _
                        "U_MISTPGRC = " & TotalGRCostActual & ", " & _
                        "U_MISTPGRQ = " & TotalGRQtyActual & ", " & _
                        "U_MISPROCS = '" & ProjectCalculationFlag & "', " & _
                        "U_MISNETST = '" & NetStatus & "', " & _
                        "U_MISHARVQ = " & ActualHarvestQty & " " & _
                        "Where DocEntry = '" & DocEntry & "' AND U_MISPROID = '" & ProjectCd & "' "

                objRecSet.DoQuery(StrSql)

            End If

        Else
            objApplication.StatusBar.SetText("Wrong Format Contact IT Support ~1.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo Setnothing
        End If

        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

        If pMode = "ADD" Then
            objApplication.StatusBar.SetText("Operation completed successfully. (Doc.Entry No.: " & Trim(CStr(intDocEntry)) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objApplication.StatusBar.SetText("Operation Update completed successfully.(Doc.Entry No.: " & Trim(CStr(DocEntry)) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        fctTestingSave = True
        pBubbleEvent = True

        GoTo Setnothing


ErrorHandler:
        If Err.Number <> 0 Then
            If oCompany.InTransaction Then
                Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            objApplication.StatusBar.SetText("Fail saving data.(" & Trim(oCompany.GetLastErrorDescription) & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

Setnothing:
        If oCompany.InTransaction Then
            Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        objRecSet = Nothing

    End Function

    Public Sub subFormLoadLookUp(ByVal pLookUpTipe As String)
        On Error GoTo ErrorHandler

        Dim intForm As Integer

        strCurntLookUp = pLookUpTipe

        Select Case pLookUpTipe
            Case "SearchNetId"

                If fctFormExist(LookUp_FormId, intForm) Then
                    objApplication.Forms.Item(intForm).Select()
                Else
                    subScrPaintLookUp(pLookUpTipe)
                    'subSetToolbar(objFormLookUp, True, _
                    '            False, False, False, False, False, _
                    '            False, False, False, False, False, False, False, False)
                End If

            Case "SearchNetIdHarvest"

                If fctFormExist(LookUpNet_FormId, intForm) Then
                    objApplication.Forms.Item(intForm).Select()
                Else
                    subScrPaintLookUpNet(pLookUpTipe)
                    'subSetToolbar(objFormLookUp, True, _
                    '            False, False, False, False, False, _
                    '            False, False, False, False, False, False, False, False)
                End If
        End Select


ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail Load LookUp !~8.0001~", vbExclamation, "SAP BO")
        End If

    End Sub

    Public Sub subFormLoadDistNumber(ByVal row As Integer)
        On Error GoTo ErrorHandler

        Dim intForm As Integer

        strCurntLookUpDistNumber = row

        If fctFormExist(LookUpDistNumber_FormId, intForm) Then
            objApplication.Forms.Item(intForm).Select()
        Else
            subScrPaintLookUpDistNumber(row)
            'subSetToolbar(objFormLookUp, True, _
            '            False, False, False, False, False, _
            '            False, False, False, False, False, False, False, False)
        End If

ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail Load DistNumber !~7.0001~", vbExclamation, "SAP BO")
        End If


    End Sub

    Public Sub subFormLoadLookUpMortal(ByVal pLookUpTipe As String)
        On Error GoTo ErrorHandler

        Dim intForm As Integer

        strCurntLookUpMortal = pLookUpTipe

        If fctFormExist(LookUpMortal_FormId, intForm) Then
            objApplication.Forms.Item(intForm).Select()
        Else
            subScrPaintLookUpMortal(pLookUpTipe)
        End If

ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail Load Look Up Mortality !~8.0001~", vbExclamation, "SAP BO")
        End If

    End Sub

    Public Sub subFormLoadLookUpBatch(ByVal pLookUpTipe As String)
        On Error GoTo ErrorHandler

        Dim intForm As Integer

        strCurntLookUpBatch = pLookUpTipe

        If fctFormExist(LookUpBatch_FormId, intForm) Then
            objApplication.Forms.Item(intForm).Select()
        Else
            subScrPaintLookUpBatch(pLookUpTipe)
            'subSetToolbar(objFormLookUp, True, _
            '            False, False, False, False, False, _
            '            False, False, False, False, False, False, False, False)
        End If

ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail Load Look Up Batch !~8.0001~", vbExclamation, "SAP BO")
        End If

    End Sub

    Public Sub subScrPaintLookUpMortal(ByVal pLookUpTipe As String)
        subScrPaint("Mortality.srf", LookUpMortal_FormId, intFormCountLookUpMortal, oFormLookUpMortal)

        subSetFirstLoadLookUpMortal(True, pLookUpTipe)

        blnModalLookUpMortal = True
    End Sub

    Public Sub subScrPaintLookUpBatch(ByVal pLookUpTipe As String)
        subScrPaint("ScrPaintLookUpBatch.srf", LookUpBatch_FormId, intFormCountLookUpBatch, oFormLookUpBatch)

        subSetFirstLoadLookUpBatch(True, pLookUpTipe)

        blnModalLookUpBatch = True
    End Sub

    Public Sub subScrPaintLookUpDistNumber(ByVal row As Integer)
        subScrPaint("ListBatch.srf", LookUpDistNumber_FormId, intFormCountLookUpDistNumber, oFormLookUpDistNumber)

        subSetFirstLoadLookUpDistNumber(True, row)

        blnModalLookUpDistNumber = True
    End Sub
    'Public Sub subFormLoadLookUpSpecies(ByVal pLookUpTipe As String)
    '    Dim intForm As Integer

    '    strCurntLookUpSpecies = pLookUpTipe

    '    If fctFormExist(LookUpSpecies_FormId, intForm) Then
    '        objApplication.Forms.Item(intForm).Select()
    '    Else
    '        subScrPaintLookUpSpecies(pLookUpTipe)
    '        'subSetToolbar(objFormLookUp, True, _
    '        '            False, False, False, False, False, _
    '        '            False, False, False, False, False, False, False, False)
    '    End If
    'End Sub

    Public Sub subScrPaintLookUp(ByVal pLookUpTipe As String)
        subScrPaint("ScrPaintLookUpNet.srf", LookUp_FormId, intFormCountLookUp, oFormLookUp)

        subSetFirstLoadLookUp(True, pLookUpTipe)

        blnModalLookUp = True
    End Sub

    Public Sub subScrPaintLookUpNet(ByVal pLookUpTipe As String)
        subScrPaint("ScrPaintLookUpNetHarvest.srf", LookUpNet_FormId, intFormCountLookUpNet, oFormLookUpNet)

        subSetFirstLoadLookUpNet(True, pLookUpTipe)

        blnModalLookUp = True
    End Sub

    'Public Sub subScrPaintLookUpSpecies(ByVal pLookUpTipe As String)
    '    subScrPaint("ScrPaintLookUpSpecies.srf", LookUpSpecies_FormId, intFormCountLookUpSpecies, oFormLookUpSpecies)

    '    subSetFirstLoadLookUpSpecies(True, pLookUpTipe)

    '    blnModalLookUp = True
    'End Sub
    Public Sub subSetFirstLoadLookUpDistNumber(ByVal pFirstLoad As Boolean, ByVal row As Integer)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        oMatrix = oFormLookUpDistNumber.Items.Item("MtxBatch").Specific
        oColumns = oMatrix.Columns

        If pFirstLoad Then
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("Batch", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("NetId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("RitNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpDistNumber.DataSources.UserDataSources.Add("BoxNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        End If

        oColumns.Item("Batch").DataBind.SetBound(True, "", "Batch")
        oColumns.Item("NetId").DataBind.SetBound(True, "", "NetId")
        oColumns.Item("DocDate").DataBind.SetBound(True, "", "DocDate")
        oColumns.Item("RitNo").DataBind.SetBound(True, "", "RitNo")
        oColumns.Item("BoxNo").DataBind.SetBound(True, "", "BoxNo")

        'If pLookUpTipe = "SearchDistNumber" Then
        'subInsertDataIntoBatch(objFormBatch, row)
        'End If


        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oMatrix = Nothing
        oColumns = Nothing
    End Sub

    Public Sub subSetFirstLoadLookUpMortal(ByVal pFirstLoad As Boolean, ByVal pLookUpTipe As String)

        If pFirstLoad Then
            oFormLookUpMortal.DataSources.UserDataSources.Add("PrjCd", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oFormLookUpMortal.DataSources.UserDataSources.Add("DtDead", SAPbouiCOM.BoDataType.dt_DATE)
            oFormLookUpMortal.DataSources.UserDataSources.Add("MortQty", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            oFormLookUpMortal.DataSources.UserDataSources.Add("Reason", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 255)
        End If

        oFormLookUpMortal.Items.Item("PrjCd").Specific.databind.setbound(True, "", "PrjCd")
        oFormLookUpMortal.Items.Item("DtDead").Specific.databind.setbound(True, "", "DtDead")
        oFormLookUpMortal.Items.Item("MortQty").Specific.databind.setbound(True, "", "MortQty")
        oFormLookUpMortal.Items.Item("Reason").Specific.databind.setbound(True, "", "Reason")

        oFormLookUpMortal.Items.Item("PrjCd").Specific.value = objFormProjectMaster.Items.Item("MISPROID").Specific.value
        oFormLookUpMortal.Items.Item("DtDead").Click()
        oFormLookUpMortal.Items.Item("PrjCd").Enabled = False

    End Sub

    Public Sub subSetFirstLoadLookUpBatch(ByVal pFirstLoad As Boolean, ByVal pLookUpTipe As String)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        oMatrix = oFormLookUpBatch.Items.Item("mtxSrcBtch").Specific
        oColumns = oMatrix.Columns

        If pFirstLoad Then
            oFormLookUpBatch.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            oFormLookUpBatch.DataSources.UserDataSources.Add("FinBatch", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpBatch.DataSources.UserDataSources.Add("Strain", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpBatch.DataSources.UserDataSources.Add("Acronym", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oFormLookUpBatch.DataSources.UserDataSources.Add("Desc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        End If

        oColumns.Item("colFinGB").DataBind.SetBound(True, "", "FinBatch")
        oColumns.Item("colStrain").DataBind.SetBound(True, "", "Strain")
        oColumns.Item("colAcro").DataBind.SetBound(True, "", "Acronym")
        oColumns.Item("colDesc").DataBind.SetBound(True, "", "Desc")

        If pLookUpTipe = "SearchBatch" Then
            subSearchData("SearchBatch")
        End If


        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oMatrix = Nothing
        oColumns = Nothing
    End Sub

    Public Sub subSetFirstLoadLookUpNet(ByVal pFirstLoad As Boolean, ByVal pLookUpTipe As String)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        Select Case pLookUpTipe
            Case "SearchNetIdHarvest"
                oMatrix = oFormLookUpNet.Items.Item("mtxSearch").Specific
                oColumns = oMatrix.Columns

                If pFirstLoad Then
                    oFormLookUpNet.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oFormLookUpNet.DataSources.UserDataSources.Add("NetCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUpNet.DataSources.UserDataSources.Add("FarmCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUpNet.DataSources.UserDataSources.Add("Region", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    'oFormLookUpNet.DataSources.UserDataSources.Add("NetLoc", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
                    oFormLookUpNet.DataSources.UserDataSources.Add("NetCap", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oFormLookUpNet.DataSources.UserDataSources.Add("NetPurps", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUpNet.DataSources.UserDataSources.Add("NetSts", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                End If

                oColumns.Item("colNetCd").DataBind.SetBound(True, "", "NetCode")
                oColumns.Item("colFarmCd").DataBind.SetBound(True, "", "FarmCode")
                oColumns.Item("colReg").DataBind.SetBound(True, "", "Region")
                'oColumns.Item("colNetLoc").DataBind.SetBound(True, "", "NetLoc")
                oColumns.Item("colNetCap").DataBind.SetBound(True, "", "NetCap")
                oColumns.Item("colNetPu").DataBind.SetBound(True, "", "NetPurps")
                oColumns.Item("ColNetSts").DataBind.SetBound(True, "", "NetSts")
        End Select

        subSearchDataNet("SearchNetIdHarvest")

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oMatrix = Nothing
        oColumns = Nothing

    End Sub

    Public Sub subSetFirstLoadLookUp(ByVal pFirstLoad As Boolean, ByVal pLookUpTipe As String)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing

        Select Case pLookUpTipe
            Case "SearchNetId"
                oMatrix = oFormLookUp.Items.Item("mtxSearch").Specific
                oColumns = oMatrix.Columns

                If pFirstLoad Then
                    oFormLookUp.DataSources.UserDataSources.Add("No", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oFormLookUp.DataSources.UserDataSources.Add("NetCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUp.DataSources.UserDataSources.Add("FarmCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUp.DataSources.UserDataSources.Add("Region", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    'oFormLookUp.DataSources.UserDataSources.Add("NetLoc", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
                    oFormLookUp.DataSources.UserDataSources.Add("NetCap", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
                    oFormLookUp.DataSources.UserDataSources.Add("NetPurps", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    oFormLookUp.DataSources.UserDataSources.Add("NetSts", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                    'oFormLookUp.DataSources.UserDataSources.Add("NetRate", SAPbouiCOM.BoDataType.dt_PERCENT)
                    'oFormLookUp.DataSources.UserDataSources.Add("NetFCR", SAPbouiCOM.BoDataType.dt_PERCENT)
                End If

                oColumns.Item("colNetCd").DataBind.SetBound(True, "", "NetCode")
                oColumns.Item("colFarmCd").DataBind.SetBound(True, "", "FarmCode")
                oColumns.Item("colReg").DataBind.SetBound(True, "", "Region")
                'oColumns.Item("colNetLoc").DataBind.SetBound(True, "", "NetLoc")
                oColumns.Item("colNetCap").DataBind.SetBound(True, "", "NetCap")
                oColumns.Item("colNetPu").DataBind.SetBound(True, "", "NetPurps")
                oColumns.Item("ColNetSts").DataBind.SetBound(True, "", "NetSts")
                'oColumns.Item("colNetRate").DataBind.SetBound(True, "", "NetRate")
                'oColumns.Item("colNetFCR").DataBind.SetBound(True, "", "NetFCR")

        End Select


        If pLookUpTipe = "SearchNetId" Then
            subSearchData("SearchNetId")
        End If


        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oMatrix = Nothing
        oColumns = Nothing
    End Sub

    Public Sub subSearchDataNet(ByVal pTable As String)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing
        Dim strSQL As String
        Dim intLoop As Integer
        Dim oUserDataSource(6) As SAPbouiCOM.UserDataSource
        Dim oRecSetHarvest As SAPbobsCOM.Recordset = Nothing

        strSQL = ""
        Select Case pTable
            Case "SearchNetIdHarvest"

                oMatrix = oFormLookUpNet.Items.Item("mtxSearch").Specific
                oColumns = oMatrix.Columns
                oUserDataSource(0) = oFormLookUpNet.DataSources.UserDataSources.Item("No")
                oUserDataSource(1) = oFormLookUpNet.DataSources.UserDataSources.Item("NetCode")
                oUserDataSource(2) = oFormLookUpNet.DataSources.UserDataSources.Item("FarmCode")
                oUserDataSource(3) = oFormLookUpNet.DataSources.UserDataSources.Item("Region")
                'oUserDataSource(4) = oFormLookUpNet.DataSources.UserDataSources.Item("NetLoc")
                oUserDataSource(4) = oFormLookUpNet.DataSources.UserDataSources.Item("NetCap")
                oUserDataSource(5) = oFormLookUpNet.DataSources.UserDataSources.Item("NetPurps")
                oUserDataSource(6) = oFormLookUpNet.DataSources.UserDataSources.Item("NetSts")

                oMatrix.Clear()

                oFormLookUpNet.Title = "List Net"
                oColumns.Item("colNetCd").TitleObject.Caption = "Net Code"
                oColumns.Item("colFarmCd").TitleObject.Caption = "Farm Code"
                oColumns.Item("colReg").TitleObject.Caption = "Region"
                'oColumns.Item("colNetLoc").TitleObject.Caption = "Net Location"
                oColumns.Item("colNetCap").TitleObject.Caption = "Net Capacity (M3)"
                oColumns.Item("colNetPu").TitleObject.Caption = "Net Purposes"
                oColumns.Item("ColNetSts").TitleObject.Caption = "Net Status"

                oRecSetHarvest = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                strSQL = "select T0.U_MISNETCD NetCode, T0.U_MISFARMC FarmCode, T0.U_MISREGCD Region, " & _
                        " T0.U_MISNETCA NetCapacity, T0.U_MISNETPU NetPurposes, T0.U_MISRECST NetStatus " & _
                        "from [@MIS_NETMS] T0 where T0.U_MISNETCD LIKE '%" & _
                         oFormProjectHarvest.Items.Item("MISNETID").Specific.String & "%'"


                oRecSetHarvest.DoQuery(strSQL)

                If oRecSetHarvest.RecordCount > 0 Then
                    intLoop = 0
                    Do While Not oRecSetHarvest.EoF
                        intLoop = intLoop + 1
                        oUserDataSource(1).Value = oRecSetHarvest.Fields.Item("NetCode").Value
                        oUserDataSource(2).Value = oRecSetHarvest.Fields.Item("FarmCode").Value
                        oUserDataSource(3).Value = oRecSetHarvest.Fields.Item("Region").Value
                        'oUserDataSource(4).Value = oRecSetHarvest.Fields.Item("NetLocation").Value
                        oUserDataSource(4).Value = oRecSetHarvest.Fields.Item("NetCapacity").Value
                        oUserDataSource(5).Value = oRecSetHarvest.Fields.Item("NetPurposes").Value
                        oUserDataSource(6).Value = oRecSetHarvest.Fields.Item("NetStatus").Value
                        oMatrix.AddRow()
                        oRecSetHarvest.MoveNext()
                    Loop
                Else
                    objApplication.StatusBar.SetText("Status Net Master Not Active ~23.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oFormLookUpNet.Close()
                    GoTo setnothing
                End If

                oColumns.Item("colNetCd").Editable = False
                oColumns.Item("colFarmCd").Editable = False
                oColumns.Item("colReg").Editable = False
                'oColumns.Item("colNetLoc").Editable = False
                oColumns.Item("colNetCap").Editable = False
                oColumns.Item("colNetPu").Editable = False
                oColumns.Item("ColNetSts").Editable = False

        End Select


setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSetHarvest)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oRecSetHarvest = Nothing
        oMatrix = Nothing
        oColumns = Nothing
        Select Case pTable
            Case "SearchNetIdHarvest"
                For intLoop = 0 To 6
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserDataSource(intLoop))
                    oUserDataSource(intLoop) = Nothing
                Next
        End Select
    End Sub

    Public Sub subSearchData(ByVal pTable As String)
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oColumns As SAPbouiCOM.Columns = Nothing
        Dim strSQL As String
        Dim intLoop As Integer
        Dim oUserDataSource(6) As SAPbouiCOM.UserDataSource
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing

        strSQL = ""
        Select Case pTable
            Case "SearchNetId"

                oMatrix = oFormLookUp.Items.Item("mtxSearch").Specific
                oColumns = oMatrix.Columns
                oUserDataSource(0) = oFormLookUp.DataSources.UserDataSources.Item("No")
                oUserDataSource(1) = oFormLookUp.DataSources.UserDataSources.Item("NetCode")
                oUserDataSource(2) = oFormLookUp.DataSources.UserDataSources.Item("FarmCode")
                oUserDataSource(3) = oFormLookUp.DataSources.UserDataSources.Item("Region")
                'oUserDataSource(4) = oFormLookUp.DataSources.UserDataSources.Item("NetLoc")
                oUserDataSource(4) = oFormLookUp.DataSources.UserDataSources.Item("NetCap")
                oUserDataSource(5) = oFormLookUp.DataSources.UserDataSources.Item("NetPurps")
                oUserDataSource(6) = oFormLookUp.DataSources.UserDataSources.Item("NetSts")
                'oUserDataSource(8) = oFormLookUp.DataSources.UserDataSources.Item("NetRate")
                'oUserDataSource(9) = oFormLookUp.DataSources.UserDataSources.Item("NetFCR")

                oMatrix.Clear()

                oFormLookUp.Title = "List Net"
                oColumns.Item("colNetCd").TitleObject.Caption = "Net Code"
                oColumns.Item("colFarmCd").TitleObject.Caption = "Farm Code"
                oColumns.Item("colReg").TitleObject.Caption = "Region"
                'oColumns.Item("colNetLoc").TitleObject.Caption = "Net Location"
                oColumns.Item("colNetCap").TitleObject.Caption = "Net Capacity (M3)"
                oColumns.Item("colNetPu").TitleObject.Caption = "Net Purposes"
                oColumns.Item("ColNetSts").TitleObject.Caption = "Net Status"
                'oColumns.Item("colNetRate").TitleObject.Caption = "Net Rate"
                'oColumns.Item("colNetFCR").TitleObject.Caption = "Feed Rate"

                oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If objFormProjectMaster.Items.Item("btnOK").Specific.caption = "FIND" Then

                    strSQL = "select T0.U_MISNETCD NetCode, T0.U_MISFARMC FarmCode, T0.U_MISREGCD Region, " & _
                            " T0.U_MISNETCA NetCapacity, T0.U_MISNETPU NetPurposes, T0.U_MISRECST NetStatus " & _
                            "from [@MIS_NETMS] T0 where T0.U_MISNETCD LIKE '%" & _
                             objFormProjectMaster.Items.Item("MISNETID").Specific.String & "%'"


                    oRecSet.DoQuery(strSQL)

                    If oRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not oRecSet.EoF
                            intLoop = intLoop + 1
                            oUserDataSource(1).Value = oRecSet.Fields.Item("NetCode").Value
                            oUserDataSource(2).Value = oRecSet.Fields.Item("FarmCode").Value
                            oUserDataSource(3).Value = oRecSet.Fields.Item("Region").Value
                            'oUserDataSource(4).Value = oRecSet.Fields.Item("NetLocation").Value
                            oUserDataSource(4).Value = oRecSet.Fields.Item("NetCapacity").Value
                            oUserDataSource(5).Value = oRecSet.Fields.Item("NetPurposes").Value
                            oUserDataSource(6).Value = oRecSet.Fields.Item("NetStatus").Value
                            'oUserDataSource(8).Value = oRecSet.Fields.Item("NetRate").Value
                            'oUserDataSource(9).Value = oRecSet.Fields.Item("NetFCR").Value
                            oMatrix.AddRow()
                            oRecSet.MoveNext()
                        Loop
                    Else
                        objApplication.StatusBar.SetText("Status Net Master Not Active ~19.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oFormLookUp.Close()
                        GoTo setnothing
                    End If

                Else
                    strSQL = "select T0.U_MISNETCD NetCode, T0.U_MISFARMC FarmCode, T0.U_MISREGCD Region, " & _
                            " T0.U_MISNETCA NetCapacity, T0.U_MISNETPU NetPurposes, T0.U_MISRECST NetStatus " & _
                            "from [@MIS_NETMS] T0 where T0.U_MISRECST = 'O' AND T0.U_MISNETCD LIKE '%" & _
                             objFormProjectMaster.Items.Item("MISNETID").Specific.String & "%'"


                    oRecSet.DoQuery(strSQL)

                    If oRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not oRecSet.EoF
                            intLoop = intLoop + 1
                            oUserDataSource(1).Value = oRecSet.Fields.Item("NetCode").Value
                            oUserDataSource(2).Value = oRecSet.Fields.Item("FarmCode").Value
                            oUserDataSource(3).Value = oRecSet.Fields.Item("Region").Value
                            'oUserDataSource(4).Value = oRecSet.Fields.Item("NetLocation").Value
                            oUserDataSource(4).Value = oRecSet.Fields.Item("NetCapacity").Value
                            oUserDataSource(5).Value = oRecSet.Fields.Item("NetPurposes").Value
                            oUserDataSource(6).Value = oRecSet.Fields.Item("NetStatus").Value
                            'oUserDataSource(8).Value = oRecSet.Fields.Item("NetRate").Value
                            'oUserDataSource(9).Value = oRecSet.Fields.Item("NetFCR").Value
                            oMatrix.AddRow()
                            oRecSet.MoveNext()
                        Loop
                    Else
                        objApplication.StatusBar.SetText("Status Net Master Not Active ~19.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oFormLookUp.Close()
                        GoTo setnothing
                    End If

                End If

                oColumns.Item("colNetCd").Editable = False
                oColumns.Item("colFarmCd").Editable = False
                oColumns.Item("colReg").Editable = False
                'oColumns.Item("colNetLoc").Editable = False
                oColumns.Item("colNetCap").Editable = False
                oColumns.Item("colNetPu").Editable = False
                oColumns.Item("ColNetSts").Editable = False


            Case "SearchBatch"

                oMatrix = oFormLookUpBatch.Items.Item("mtxSrcBtch").Specific
                oColumns = oMatrix.Columns
                oUserDataSource(0) = oFormLookUpBatch.DataSources.UserDataSources.Item("No")
                oUserDataSource(1) = oFormLookUpBatch.DataSources.UserDataSources.Item("FinBatch")
                oUserDataSource(2) = oFormLookUpBatch.DataSources.UserDataSources.Item("Strain")
                oUserDataSource(3) = oFormLookUpBatch.DataSources.UserDataSources.Item("Acronym")
                oUserDataSource(4) = oFormLookUpBatch.DataSources.UserDataSources.Item("Desc")

                oMatrix.Clear()

                oFormLookUpBatch.Title = "List Fingerling Batch"

                oColumns.Item("colFinGB").TitleObject.Caption = "Fingerling Batch Code"
                oColumns.Item("colStrain").TitleObject.Caption = "Strain"
                oColumns.Item("colAcro").TitleObject.Caption = "Acronym"
                oColumns.Item("colDesc").TitleObject.Caption = "Desc"

                strSQL = "select T0.U_MISFINGB FingerlingBatch, T0.U_MISSTRN Strain, T0.U_MISACRO Acronym, " & _
                        "T0.U_MISDESC Descrip " & _
                        "from [@MIS_FINGB] T0 where T0.U_MISFINGB LIKE '%" & _
                         objFormProjectMaster.Items.Item("MISGENCD").Specific.String & "%'"

                oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecSet.DoQuery(strSQL)

                If oRecSet.RecordCount > 0 Then
                    intLoop = 0
                    Do While Not oRecSet.EoF
                        intLoop = intLoop + 1
                        oUserDataSource(1).Value = oRecSet.Fields.Item("FingerlingBatch").Value
                        oUserDataSource(2).Value = oRecSet.Fields.Item("Strain").Value
                        oUserDataSource(3).Value = oRecSet.Fields.Item("Acronym").Value
                        oUserDataSource(4).Value = oRecSet.Fields.Item("Descrip").Value
                        oMatrix.AddRow()
                        oRecSet.MoveNext()
                    Loop
                Else
                    objApplication.StatusBar.SetText("No Data ~19.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oFormLookUpBatch.Close()
                    GoTo setnothing
                End If

                oColumns.Item("colFinGB").Editable = False
                oColumns.Item("colStrain").Editable = False
                oColumns.Item("colAcro").Editable = False
                oColumns.Item("colDesc").Editable = False

        End Select
        GoTo setnothing

setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
        oRecSet = Nothing
        oMatrix = Nothing
        oColumns = Nothing
        Select Case pTable
            Case "SearchNetId"
                For intLoop = 0 To 6
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserDataSource(intLoop))
                    oUserDataSource(intLoop) = Nothing
                Next
            Case "SearchBatch"
                For intLoop = 0 To 4
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserDataSource(intLoop))
                    oUserDataSource(intLoop) = Nothing
                Next
        End Select

    End Sub

    Private Sub subValidateFarmcode(ByVal FarmCode As String, ByVal Species As String, ByVal Klasifikasi As String)
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim strSql As String

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strSql = "Select U_MISSURVR SurvivalRate, U_MISFCR FCR, U_MISHARVP DayCul From [@MIS_RATE]  where U_MISFARMC = '" & FarmCode & "' AND U_MISFSPEC = " & Species & " AND U_MISHATGRO = '" & Klasifikasi & "' "

        objRecSet.DoQuery(strSql)

        If objRecSet.RecordCount > 0 Then
            objFormProjectMaster.Items.Item("MISESTLF").Specific.string = objRecSet.Fields.Item("SurvivalRate").Value
            objFormProjectMaster.Items.Item("MISFCR").Specific.string = objRecSet.Fields.Item("FCR").Value
            objFormProjectMaster.Items.Item("MISHARVP").Specific.string = objRecSet.Fields.Item("DayCul").Value
        Else
            objApplication.MessageBox("You Must Fill Rate Master", 1, "OK")

            'objFormProjectMaster.Items.Item("MISESTLF").Specific.string = 100
            'objFormProjectMaster.Items.Item("MISFCR").Specific.string = 0
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
        objRecSet = Nothing
    End Sub

    '    Private Sub subInsertDataIntoBatch(ByVal pform As SAPbouiCOM.Form, ByVal row As Integer)
    '        On Error GoTo errorhandler
    '        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
    '        Dim objMatrixBatch As SAPbouiCOM.Matrix = Nothing
    '        Dim objColumnsBatch As SAPbouiCOM.Columns = Nothing
    '        Dim TransactionType As Integer
    '        Dim StrSqlHeader As String
    '        Dim StrSqlDetail As String
    '        Dim intLoop As Integer

    '        'TransactionType = objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string
    '        'If TransactionType = 2 Or TransactionType = 3 Then
    '        '    Select TransactionType
    '        '        Case 2
    '        objMatrixBatch = oFormLookUpDistNumber.Items.Item("MtxBatch").Specific
    '        objColumnsBatch = objMatrixBatch.Columns

    '        'objMatrixBatch = objFormBatch.Items.Item("3").Specific
    '        'objColumnsBatch = objMatrixBatch.Columns


    '        objMatrixBatch.Clear()
    '        objMatrixBatch.AddRow()

    '        StrSqlDetail = " SELECT DistNumber Batch, T1.U_MISNETID NetId, T0.DocDate DocDate, T0.U_MISRITNO RitNo, T1.U_MISBOXNO BoxNo " & _
    '                         "FROM OIGE T0 " & _
    '                         "INNER JOIN IGE1 T1 " & _
    '                         "ON T0.DocEntry = T1.DocEntry " & _
    '                         "INNER JOIN OBTN T2 " & _
    '                         "ON T1.itemcode = T2.ItemCode " & _
    '                        "AND T1.DocDate = T2.InDate " & _
    '                        "AND T1.U_MISBOXNO = Right(T2.DistNumber,2) " & _
    '                        "AND T0.U_MISRITNO = SUBSTRING(T2.DistNumber,17,2) " & _
    '                       "Where T0.DocEntry = " & DocEntry & " "
    '        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        objRecSet.DoQuery(StrSqlDetail)

    '        'BoxNo = objRecSet.Fields.Item("BoxNo")

    '        If objRecSet.RecordCount > 0 Then
    '            intLoop = 0
    '            Do While Not objRecSet.EoF
    '                intLoop = intLoop + 1
    '                'objColumnsBatch.Item("2").Cells.Item(row).Specific.string = objRecSet.Fields.Item("Batch").Value
    '                objColumnsBatch.Item("Batch").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Batch").Value
    '                objColumnsBatch.Item("NetId").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("NetId").Value
    '                objColumnsBatch.Item("DocDate").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("DocDate").Value
    '                objColumnsBatch.Item("RitNo").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("RitNo").Value
    '                objColumnsBatch.Item("BoxNo").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
    '                objMatrixBatch.AddRow()
    '                objRecSet.MoveNext()
    '            Loop
    '            'oFormLookUpDistNumber.Close()
    '        Else
    '            objApplication.StatusBar.SetText("No Data", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            oFormLookUpDistNumber.Close()
    '            GoTo errorhandler
    '        End If
    '        '    End Select
    '        'End If


    'errorhandler:

    '        'If TransactionType = 2 Or TransactionType = 3 Then
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrixBatch)
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumnsBatch)
    '        ' End If
    '        objRecSet = Nothing
    '        objMatrixBatch = Nothing
    '        objColumnsBatch = Nothing

    '    End Sub

    Private Sub subInsertDataIntoInvTransfer(ByRef DocEntry As Integer, ByRef RowNo As Integer, ByVal pform As SAPbouiCOM.Form, ByVal pVal As SAPbouiCOM.ItemEvent)
        On Error GoTo errorhandler
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim TransactionType As Integer
        Dim StrSqlHeader As String
        Dim StrSqlDetail As String
        Dim intLoop As Integer
        Dim intMsg As Integer

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        objFormListInvTransfer.Close()
        TransactionType = objFormInvTransferUDF.Items.Item("U_MISTRXTP").Specific.string

        'If TransactionType >= 1 And TransactionType <= 5 Then
        If TransactionType = 2 Then
            Select Case TransactionType
                Case 2
                    'HEADER
                    StrSqlHeader = "Select T0.DocEntry DocEntry, T0.DocNum DocNum, T0.U_MISPRNO PRNO, T0.U_MISREQBY RequestBy, T0.Filler ToWarehouse " & _
                                   "From OWTR T0 INNER JOIN WTR1 T1 ON T0.Docentry = T1.Docentry where T0.DocEntry = " & DocEntry & " "
                    objRecSet.DoQuery(StrSqlHeader)

                    If objRecSet.RecordCount > 0 Then
                        objFormInvTransfer.Items.Item("18").Specific.string = objRecSet.Fields.Item("ToWarehouse").Value
                        objFormInvTransferUDF.Items.Item("U_MISPRNO").Specific.string = objRecSet.Fields.Item("PRNO").Value
                        objFormInvTransferUDF.Items.Item("U_MISREQBY").Specific.string = objRecSet.Fields.Item("RequestBy").Value
                    Else
                        objApplication.StatusBar.SetText("No Data ~15.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                    'DETAIL
                    objMatrix = objFormInvTransfer.Items.Item("23").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    intMsg = objApplication.MessageBox("Do You Want To Copy All Row", 1, "YES", "NO")
                    If intMsg = 1 Then
                        StrSqlDetail = "Select T1.ItemCode ItemCode,T1.Dscription Description, T1.Quantity Quantity, T1.unitMsr UnitMeasure " & _
                                        "From OWTR T0 INNER JOIN WTR1 T1 ON T0.DOCENTRY = T1.DOCENTRY " & _
                                        "Where T0.DocEntry = " & DocEntry & ""

                        objRecSet.DoQuery(StrSqlDetail)
                    Else
                        StrSqlDetail = "Select T1.ItemCode ItemCode,T1.Dscription Description, T1.Quantity Quantity, T1.unitMsr UnitMeasure " & _
                                        "From OWTR T0 INNER JOIN WTR1 T1 ON T0.DOCENTRY = T1.DOCENTRY " & _
                                        "Where T0.DocEntry = " & DocEntry & " AND T1.VisOrder = " & RowNo & " "

                        objRecSet.DoQuery(StrSqlDetail)
                    End If

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            objColumns.Item("2").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Description").Value
                            objColumns.Item("10").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            'objColumns.Item("1002").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("UnitMeasure").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop

                    Else
                        objApplication.StatusBar.SetText("No Data ~15.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGI.Close()
                        GoTo errorhandler
                    End If

            End Select
        Else
            objFormListGI.Close()
            objApplication.StatusBar.SetText("U Must Input Manual ~15.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo errorhandler
        End If

errorhandler:

        If TransactionType = 2 Or TransactionType = 3 Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
        End If
        objRecSet = Nothing
        objMatrix = Nothing
        objColumns = Nothing

    End Sub

    Private Sub subInsertDataIntoGoodReceipt(ByRef DocEntry As Integer, ByVal pform As SAPbouiCOM.Form, ByVal pVal As SAPbouiCOM.ItemEvent)
        On Error GoTo errorhandler
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim TransactionType As Integer
        Dim StrSqlHeader As String
        Dim StrSqlDetail As String
        Dim intLoop As Integer

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        objFormListGI.Close()
        TransactionType = objFormGoodReceiptUDF.Items.Item("U_MISTRXTP").Specific.string

        'If TransactionType >= 1 And TransactionType <= 5 Then
        If TransactionType = 2 Or TransactionType = 3 Or TransactionType = 4 Then
            'objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Visible = False

            'objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = True
            ''objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Enabled = True
            'objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string = ""
            'objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Visible = False

            Select Case TransactionType
                Case 2
                    objFormGoodReceiptUDF.Freeze(True)
                    'HEADER
                    StrSqlHeader = "Select DocEntry DocEntry, DocNum DocNum, U_MISDESTW Destination, U_MISDRVNM DriverName, U_MISASDRV AssDriver, U_MISLICNO LicenceNo, U_MISSPVID SupervisorId, " & _
                                   "U_MISRITNO RitNo, U_MISDELTM Delivery From OIGE where DocEntry=" & DocEntry & ""
                    objRecSet.DoQuery(StrSqlHeader)

                    'objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = objRecSet.Fields.Item("Destination").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Specific.string = objRecSet.Fields.Item("DriverName").Value

                    objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Specific.string = objRecSet.Fields.Item("AssDriver").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Enabled = True

                    objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Specific.string = objRecSet.Fields.Item("LicenceNo").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Enabled = True

                    objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Specific.string = objRecSet.Fields.Item("SupervisorId").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Enabled = True
                    Dim Ritno As String
                    Ritno = objRecSet.Fields.Item("RitNo").Value

                    objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.select(RitNo, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Enabled = True

                    objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.string = objRecSet.Fields.Item("Delivery").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Enabled = True

                    objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Enabled = True

                    objFormGoodReceipt.Items.Item("txtBatch").Specific.string = ""
                    objFormGoodReceipt.Items.Item("txtBatch").Specific.string = objRecSet.Fields.Item("DocEntry").Value

                    objFormGoodReceiptUDF.Freeze(False)
                    objFormGoodReceipt.Freeze(True)
                    'DETAIL
                    'objMatrixBatch = objFormBatch.Items.Item("3").Specific
                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific
                    'objColumnsBatch = objMatrixBatch.Columns
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()
                    'objMatrixBatch.Clear()
                    'objMatrixBatch.AddRow()

                    'StrSqlDetail = "Select T0.ItemCode ItemCode,T0.Quantity Quantity, T0.U_MISGISQK QtyInKg, T0.U_MISGISQP QtyInPcs, " & _
                    '                "(T0.StockPrice) UnitPrice, (T0.Quantity * T0.StockPrice) Total, T1.U_MISDESTW Destination, T2.FormatCode InvOffset, T0.U_MISFISHQ FishQty, T0.U_MISBOXNO BoxNo, T0.U_MISNOSEG SegelNo " & _
                    '                "From IGE1 T0 INNER JOIN OIGE T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OACT T2 ON T0.acctCode = T2.AcctCode " & _
                    '                "Where T0.DocEntry = " & DocEntry & ""
                    StrSqlDetail = "Select T0.ItemCode ItemCode,T3.Quantity Quantity,  T3.BatchNum Batch, " & _
                    "(T0.StockPrice) UnitPrice, (T3.Quantity * T0.StockPrice) Total, T1.U_MISDESTW Destination, T2.FormatCode InvOffset, T0.U_MISFISHQ FishQty, T0.U_MISBOXNO BoxNo, T0.U_MISNOSEG SegelNo " & _
                    "From IGE1 T0 INNER JOIN OIGE T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OACT T2 ON T0.acctCode = T2.AcctCode " & _
                    "INNER JOIN IBT1 T3 ON T0.ItemCode = T3.ItemCode AND T0.WhsCode = T3.WhsCode " & _
                    "Where T3.BaseEntry = T0.DocEntry  AND T0.LineNum = T3.BaseLinNum And T3.BaseType = 60 And T0.DocEntry = " & DocEntry & " "

                    '"Select ItemCode ItemCode, U_MISBOXNO BoxNo, " & _
                    '                                   "U_MISGISQK QtyInKg, U_MISGISQP QtyInPcs, (LineTotal/Quantity) UnitPrice, LineTotal Total From IGE1 " & _
                    '                                   "Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    objColumns.Item("U_MISFISHQ").Visible = True
                    objColumns.Item("10").Visible = True
                    objColumns.Item("14").Visible = True

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            'objColumns.Item("U_MISNOSEG").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("SegelNo").Value
                            objColumns.Item("U_MISBATCH").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Batch").Value
                            'objColumns.Item("U_MISGISQK").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("QtyInKg").Value
                            'objColumns.Item("U_MISGISQP").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("QtyInPcs").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("10").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("UnitPrice").Value
                            objColumns.Item("14").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Total").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Destination").Value
                            objColumns.Item("57").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("InvOffset").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objMatrix.DeleteRow(intLoop + 1)

                    Else
                        objApplication.StatusBar.SetText("No Data ~14.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGI.Close()
                        GoTo errorhandler
                    End If

                    objColumns.Item("57").Cells.Item(objMatrix.VisualRowCount).Specific.string = ""

                    'objColumns.Item("U_MISFISHQ").Visible = False
                    'objColumns.Item("10").Visible = False
                    'objColumns.Item("14").Visible = False
                    objFormGoodReceipt.Freeze(False)

                Case 3
                    Dim RitNo As String
                    'HEADER
                    StrSqlHeader = "Select DocEntry DocEntry, DocNum ReferDoc,U_MISDELTM DeliverTime, U_MISDESTW WarehouseDest, U_MISDRVNM DriverName, U_MISASDRV AssDriver, U_MISLICNO LicenceNo, U_MISSPVID SupervisorId, " & _
                                   "U_MISRITNO RitNo From OIGE where DocEntry=" & DocEntry & ""
                    objRecSet.DoQuery(StrSqlHeader)

                    RitNo = objRecSet.Fields.Item("RitNo").Value

                    objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("ReferDoc").Value
                    'objFormGoodReceiptUDF.Items.Item("U_MISTRXWH").Specific.string = objRecSet.Fields.Item("WarehouseDest").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISDRVNM").Specific.string = objRecSet.Fields.Item("DriverName").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISASDRV").Specific.string = objRecSet.Fields.Item("AssDriver").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISLICNO").Specific.string = objRecSet.Fields.Item("LicenceNo").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISSPVID").Specific.string = objRecSet.Fields.Item("SupervisorId").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISRITNO").Specific.select(RitNo, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    objFormGoodReceiptUDF.Items.Item("U_MISDELTM").Specific.string = objRecSet.Fields.Item("DeliverTime").Value
                    objFormGoodReceipt.Items.Item("txtBatch").Specific.string = objRecSet.Fields.Item("DocEntry").Value
                    'DETAIL
                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific

                    'objMatrixBatch = objFormBatch.Items.Item("3").Specific
                    objColumns = objMatrix.Columns
                    ' objColumnsBatch = objMatrixBatch.Columns

                    objMatrix.Clear()
                    ''objMatrixBatch.Clear()
                    objMatrix.AddRow()
                    'objMatrixBatch.AddRow()

                    'StrSqlDetail = " SELECT T1.ItemCode ItemCode, T1.Quantity Quantity, T3.BatchNum Batch, T1.Quantity QtyInKg, " & _
                    '                " T1.U_MISFISHQ QtyInPcs, (T1.StockPrice) UnitPrice, (T1.Quantity * T1.StockPrice) Total, " & _
                    '                " T0.U_MISDESTW Destination , T2.FormatCode InvOffset, T1.U_MISBOXNO BoxNo, T1.U_MISNOSEG SegelNo, T1.U_MISFISHQ FishQty " & _
                    '                " From OIGE T0 INNER JOIN IGE1 T1 " & _
                    '                " ON T0.DocEntry = T1.DocEntry INNER JOIN OACT T2 " & _
                    '                " ON T1.acctCode = T2.AcctCode " & _
                    '                " Where T0.DocEntry = " & DocEntry & ""
                    'StrSqlDetail = " SELECT T1.U_MISPROID Project, T1.ItemCode ItemCode, SUM(T3.Quantity) Quantity, LEFT(T3.BatchNum,20) Batch, " & _
                    '                " SUM(T1.StockPrice) UnitPrice, SUM(T3.Quantity * T1.StockPrice) Total, " & _
                    '                " T0.U_MISDESTW Destination , T2.FormatCode InvOffset, T1.U_MISNOSEG SegelNo, SUM(T1.U_MISFISHQ) FishQty " & _
                    '                " From OIGE T0 INNER JOIN IGE1 T1 " & _
                    '                " ON T0.DocEntry = T1.DocEntry INNER JOIN OACT T2 " & _
                    '                " ON T1.acctCode = T2.AcctCode " & _
                    '                "INNER JOIN IBT1 T3 ON T1.ItemCode = T3.ItemCode AND T1.WhsCode = T3.WhsCode " & _
                    '                "Where T3.BaseEntry = T0.DocEntry AND T1.LineNum = T3.BaseLinNum And T3.BaseType = 60 And T0.DocEntry = " & DocEntry & "" & _
                    '                "GROUP BY T1.U_MISPROID, T1.ItemCode, LEFT(T3.BatchNum,20) , T1.Quantity, T1.U_MISFISHQ, " & _
                    '                "T0.U_MISDESTW, T2.FormatCode, T1.U_MISNOSEG, T1.U_MISFISHQ "

                    StrSqlDetail = "SELECT E1.U_MISPROID Project,  E1.ItemCode, SUM(E1.Quantity) Quantity, E1.U_MISPROID + '.' + E0.U_MISRITNO Batch, " & _
                                    "SUM(E1.StockPrice) UnitPrice, E4.Debit Total, E0.U_MISDESTW Destination , E2.FormatCode InvOffset, E1.U_MISNOSEG SegelNo, " & _
                                    "SUM(E1.U_MISFISHQ) FishQty FROM OIGE E0 INNER JOIN IGE1 E1 ON E0.DocEntry = E1.DocEntry INNER JOIN OACT E2 " & _
                                    "ON E1.AcctCode = E2.AcctCode LEFT JOIN OJDT E3 ON E0.DocNum = E3.BaseRef AND E0.ObjType = E3.TransType INNER JOIN JDT1 E4 " & _
                                    "ON E3.TransId = E4.TransId AND E2.AcctCode = E4.Account WHERE(E0.U_MISTRXTP = 5) AND NOT EXISTS(SELECT T1.U_MISTRXTP, T1.U_MISTRXNM, T1.DocNum, T1.DocDate, T1.CreateDate, " & _
                                    "T1.Ref1, T1.Ref2, T1.U_MISTRXWH, T1.U_MISDESTW FROM OIGN T1 WHERE(T1.U_MISTRXTP = 3 And E0.DocNum = T1.U_MISREFFD)) " & _
                                    "AND E0.DocEntry = " & DocEntry & " GROUP BY E0.U_MISTRXTP,E1.U_MISBATCH, E1.U_MISNOSEG, E0.U_MISTRXNM, E0.Docnum, E0.U_MISREFFD, E0.DocDate, " & _
                                    "E0.CreateDate, E0.U_MISTRXWH, E0.ref2, E0.Comments, E0.U_MISRITNO, E0.U_MISDRVNM, E0.U_MISASDRV, E0.U_MISSPVID, " & _
                                    "E0.U_MISDESTW, E1.U_MISPROID, E1.ItemCode, E1.Dscription, E2.FormatCode, E2.AcctName, E4.Debit "





                    objRecSet.DoQuery(StrSqlDetail)

                    objFormGoodReceiptUDF.Items.Item("U_MISGISQP").Specific.string = objRecSet.Fields.Item("FishQty").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISGISQK").Specific.string = objRecSet.Fields.Item("Quantity").Value
                    objFormGoodReceiptUDF.Items.Item("U_MISGISTV").Specific.string = objRecSet.Fields.Item("Total").Value

                    objFormGoodReceipt.Items.Item("11").Specific.string = objRecSet.Fields.Item("Project").Value

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            'objColumns.Item("U_MISNOSEG").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("SegelNo").Value
                            'objColumns.Item("U_MISGISQP").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            objColumns.Item("U_MISBATCH").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Batch").Value
                            'objColumns.Item("U_MISGISQK").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            'objColumns.Item("10").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("UnitPrice").Value
                            'objColumns.Item("U_MISGISTV").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Total").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Destination").Value
                            objColumns.Item("57").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("InvOffset").Value
                            'objColumns.Item("U_MISPROID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Project").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()

                        Loop
                        objMatrix.DeleteRow(intLoop + 1)

                    Else
                        objApplication.StatusBar.SetText("No Data ~14.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGI.Close()
                        GoTo errorhandler
                    End If

                    objColumns.Item("57").Cells.Item(objMatrix.VisualRowCount).Specific.string = ""
                    objFormGoodReceiptUDF.Items.Item("U_MISINFO").Specific.select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                Case 4
                    'HEADER
                    StrSqlHeader = "Select DocEntry DocEntry, DocNum ReferDoc " & _
                                   "From OIGE where DocEntry = " & DocEntry & ""
                    objRecSet.DoQuery(StrSqlHeader)

                    objFormGoodReceiptUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("ReferDoc").Value


                    'DETAIL
                    objMatrix = objFormGoodReceipt.Items.Item("13").Specific

                    'objMatrixBatch = objFormBatch.Items.Item("3").Specific
                    objColumns = objMatrix.Columns
                    ' objColumnsBatch = objMatrixBatch.Columns

                    objMatrix.Clear()

                    ''objMatrixBatch.Clear()
                    objMatrix.AddRow()
                    'objMatrixBatch.AddRow()
                    StrSqlDetail = " SELECT TOP 1 T1.U_MISPROID Project " & _
                                    " From OIGE T0 INNER JOIN IGE1 T1 " & _
                                    " ON T0.DocEntry = T1.DocEntry " & _
                                    "Where T0.DocEntry = " & DocEntry & " "

                    objRecSet.DoQuery(StrSqlDetail)

                    objFormGoodReceipt.Items.Item("11").Specific.string = objRecSet.Fields.Item("Project").Value


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            End Select
        Else
            objFormListGI.Close()
            objApplication.StatusBar.SetText("U Must Input Manual ~14.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo errorhandler
        End If

errorhandler:

        If TransactionType = 2 Or TransactionType = 3 Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
        End If
        objRecSet = Nothing
        objMatrix = Nothing
        objColumns = Nothing

    End Sub

    Private Sub subInsertDataIntoGoodIssue(ByRef DocEntry As Integer, ByVal pform As SAPbouiCOM.Form, ByVal pVal As SAPbouiCOM.ItemEvent)
        On Error GoTo errorhandler
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim objMatrix As SAPbouiCOM.Matrix = Nothing
        Dim objColumns As SAPbouiCOM.Columns = Nothing
        Dim TransactionType As Integer
        Dim StrSqlHeader As String
        Dim StrSqlDetail As String
        Dim intLoop As Integer

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        TransactionType = objFormGoodIssueUDF.Items.Item("U_MISTRXTP").Specific.string

        If TransactionType >= 1 And TransactionType <= 6 Then
            Select Case TransactionType
                Case 1
                    'HEADER
                    'StrSqlHeader = "Select DocNum DocNum From OIGN where DocEntry=" & DocEntry & ""
                    'objRecSet.DoQuery(StrSqlHeader)

                    'objFormGoodIssueUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value

                    'DETAIL

                    objFormGoodIssue.Freeze(True)
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    'objMatrix.AddRow()

                    StrSqlDetail = "Select ItemCode ItemCode, Quantity Quantity, " & _
                                   "U_MISFISHQ FishQty, WhsCode Warehouse, U_MISNETID NetId From IGN1 " & _
                                   "Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            objColumns.Item("U_MISNETID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("NetId").Value
                        Loop
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If

                    objFormGoodIssue.Freeze(False)

                Case 2
                    'HEADER
                    'StrSqlHeader = "Select DocNum DocNum From OIGN where DocEntry=" & DocEntry & ""
                    'objRecSet.DoQuery(StrSqlHeader)

                    'objFormGoodIssueUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value
                    objFormGoodIssue.Freeze(True)
                    'DETAIL
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    StrSqlDetail = "Select ItemCode ItemCode, Quantity Quantity, " & _
                                   "WhsCode Warehouse, U_MISNETID NetId From IGN1 " & _
                                   "Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            'objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("QtyInPcs").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            objColumns.Item("U_MISNETID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("NetId").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If
                    objFormGoodIssue.Freeze(False)
                Case 3
                    'HEADER
                    'StrSqlHeader = "Select U_MISDRVNM DriverName, U_MISASDRV AstDriver From OIGN where DocEntry=" & DocEntry & ""
                    'objRecSet.DoQuery(StrSqlHeader)

                    'objFormGoodIssueUDF.Items.Item("U_MISDRVNM").Specific.string = objRecSet.Fields.Item("DriverName").Value
                    'objFormGoodIssueUDF.Items.Item("U_MISASDRV").Specific.string = objRecSet.Fields.Item("AstDriver").Value
                    'objFormGoodIssueUDF.Items.Item("U_MISASDRV").Specific.string = objRecSet.Fields.Item("AstDriver").Value

                    objFormGoodIssue.Freeze(True)
                    'DETAIL
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    StrSqlDetail = "Select ItemCode ItemCode, Quantity Quantity, " & _
                                   "WhsCode Warehouse, U_MISFISHQ FishQty, U_MISBOXNO BoxNo From IGN1 " & _
                                   "Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            'objColumns.Item("U_MISNETID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("NetId").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If
                    objFormGoodIssue.Freeze(False)
                Case 4
                    'HEADER
                    'StrSqlHeader = "Select DocNum DocNum From OIGN where DocEntry=" & DocEntry & ""p
                    'objRecSet.DoQuery(StrSqlHeader)

                    'objFormGoodIssueUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value
                    objFormGoodIssue.Freeze(True)
                    'DETAIL
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    StrSqlDetail = "Select ItemCode ItemCode, Quantity Quantity, " & _
                                   "WhsCode Warehouse, U_MISFISHQ FishQty From IGN1 " & _
                                   "Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("BoxNo").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            'objColumns.Item("U_MISNETID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("NetId").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0004~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If
                    objFormGoodIssue.Freeze(False)
                Case 5
                    Dim Rit As String
                    Dim Boxno As String
                    'HEADER()
                    objFormGoodIssueUDF.Freeze(True)
                    StrSqlHeader = "Select U_MISRITNO RitNo, U_MISTRXWH Warehouse, DocNum DocNum From OIGN where DocEntry=" & DocEntry & ""
                    objRecSet.DoQuery(StrSqlHeader)

                    Rit = objRecSet.Fields.Item("RitNo").Value

                    objFormGoodIssueUDF.Items.Item("U_MISRITNO").Specific.Select(Rit, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    objFormGoodIssueUDF.Items.Item("U_MISTRXWH").Specific.string = objRecSet.Fields.Item("Warehouse").Value
                    objFormGoodIssueUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value

                    objFormGoodIssueUDF.Freeze(False)
                    objFormGoodIssue.Freeze(True)
                    'DETAIL
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    StrSqlDetail = "Select T0.U_MISPROID Project, T0.ItemCode ItemCode, T0.WhsCode Warehouse, T0.Quantity Quantity, " & _
                                   "T0.U_MISFISHQ FishQty, T0.U_MISBOXNO BoxNo, T0.U_MISNOSEG SealNo , T1.U_GIT InvOffset " & _
                                   "From IGN1 T0 INNER JOIN OITW T1 ON T0.ItemCode = T1.ItemCode AND T0.WhsCode = T1.WhsCode Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    Boxno = objRecSet.Fields.Item("BoxNo").Value

                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            objColumns.Item("57").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("InvOffset").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            'objColumns.Item("U_MISBOXNO").Cells.Item(intLoop).Specific.Select(Boxno, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            'objColumns.Item("U_MISNOSEG").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("SealNo").Value
                            objColumns.Item("U_MISPROID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Project").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objMatrix.DeleteRow(intLoop + 1)
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0005~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If
                    objFormGoodIssue.Freeze(False)
                Case 6
                    objFormGoodIssueUDF.Freeze(True)
                    'HEADER
                    StrSqlHeader = "Select DocNum DocNum,Comments Remark From OIGN where DocEntry=" & DocEntry & ""
                    objRecSet.DoQuery(StrSqlHeader)

                    objFormGoodIssueUDF.Items.Item("U_MISREFFD").Specific.string = objRecSet.Fields.Item("DocNum").Value
                    objFormGoodIssue.Items.Item("11").Specific.string = objRecSet.Fields.Item("Remark").Value

                    objFormGoodIssueUDF.Freeze(False)
                    objFormGoodIssue.Freeze(True)
                    'DETAIL
                    objMatrix = objFormGoodIssue.Items.Item("13").Specific
                    objColumns = objMatrix.Columns

                    objMatrix.Clear()
                    objMatrix.AddRow()

                    StrSqlDetail = "Select T0.U_MISPROID Project, T0.ItemCode ItemCode, T0.Quantity Quantity, " & _
                                   "T0.U_MISFISHQ FishQty, T0.U_MISINFO Category, " & _
                                   "T0.U_MISBOXNO BoxNo, T0.LineTotal TotalCost, T0.whsCode Warehouse, T1.U_WIP InvOffset From IGN1 T0" & _
                                   " INNER JOIN OITW T1 ON T0.ItemCode = T1.ItemCode AND T0.WhsCode = T1.WhsCode Where DocEntry = " & DocEntry & ""

                    objRecSet.DoQuery(StrSqlDetail)

                    Dim Category As String



                    If objRecSet.RecordCount > 0 Then
                        intLoop = 0
                        Do While Not objRecSet.EoF
                            intLoop = intLoop + 1
                            objColumns.Item("1").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                            objColumns.Item("9").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Quantity").Value
                            objColumns.Item("U_MISFISHQ").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("FishQty").Value
                            Category = objRecSet.Fields.Item("Category").Value
                            objColumns.Item("U_MISINFO").Cells.Item(intLoop).Specific.select(Category, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objColumns.Item("14").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("TotalCost").Value
                            objColumns.Item("15").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Warehouse").Value
                            objColumns.Item("57").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("InvOffset").Value
                            objColumns.Item("U_MISPROID").Cells.Item(intLoop).Specific.string = objRecSet.Fields.Item("Project").Value
                            objMatrix.AddRow()
                            objRecSet.MoveNext()
                        Loop
                        objMatrix.DeleteRow(intLoop + 1)
                        objFormListGR.Close()
                    Else
                        objApplication.StatusBar.SetText("No Data ~13.0006~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objFormListGR.Close()
                        GoTo errorhandler
                    End If
                    objFormGoodIssue.Freeze(False)
            End Select
        Else
            objFormListGR.Close()
            objApplication.StatusBar.SetText("U Must Input Manual ~13.0007~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            GoTo errorhandler
        End If

errorhandler:

        If TransactionType >= 1 And TransactionType <= 6 Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objMatrix)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objColumns)
        End If
        objRecSet = Nothing
        objMatrix = Nothing
        objColumns = Nothing

    End Sub

    Private Sub subItemMasterAddObject(ByRef pForm As SAPbouiCOM.Form, ByRef pFirstLoad As Boolean)
        Dim objItem As SAPbouiCOM.Item

        'add button
        objItem = pForm.Items.Add("btnGenItm", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 471
        objItem.Height = 19
        objItem.Left = 150
        objItem.Width = 150
        objItem.Specific.Caption = "Generated Item Code"

        objItem = Nothing
    End Sub

    Private Sub subInvTransferAddObject(ByRef pForm As SAPbouiCOM.Form, ByRef pFirstLoad As Boolean)
        Dim objItem As SAPbouiCOM.Item

        'add button
        objItem = pForm.Items.Add("btnTrnsfr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 335
        objItem.Height = 19
        objItem.Left = 150
        objItem.Width = 150
        objItem.Specific.Caption = "Copy From Transfer In"

        objItem = Nothing
    End Sub

    Private Sub subGoodIssueAddObject(ByRef pForm As SAPbouiCOM.Form, ByRef pFirstLoad As Boolean)
        Dim objItem As SAPbouiCOM.Item

        'add button
        objItem = pForm.Items.Add("btnCopyGR", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 150
        objItem.Width = 120
        objItem.Specific.Caption = "Copy From Good Receipt"

        objItem = pForm.Items.Add("btnGen", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 150
        objItem.Width = 80
        objItem.Specific.Caption = "Upload"

        objItem = pForm.Items.Add("btnCalc", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 380
        objItem.Width = 100
        objItem.Specific.Caption = "Calculate Qty"

        'objItem = pForm.Items.Add("txtGen", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        'objItem.Top = 318
        'objItem.Height = 19
        'objItem.Left = 240
        'objItem.Width = 150

        pForm.Items.Item("btnCopyGR").Visible = False
        pForm.Items.Item("btnGen").Visible = False

        'pForm.Items.Item("txtGen").Visible = False

        objItem = Nothing
    End Sub

    Private Sub subBatchAddObject(ByRef pForm As SAPbouiCOM.Form, ByRef pFirstLoad As Boolean)
        Dim objItem As SAPbouiCOM.Item

        ' add text
        objItem = pForm.Items.Add("txtBatch", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        objItem.Top = 4
        objItem.Height = 21
        objItem.Left = 120
        objItem.Width = 300
        objItem.TextStyle = 4
        objItem.Specific.Caption = "Double Click Rows from Documents To Generate Batch Number "

        ''add button
        'objItem = pForm.Items.Add("btnBatch", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        'objItem.Top = 354
        'objItem.Height = 21
        'objItem.Left = 190
        'objItem.Width = 150
        'objItem.Specific.Caption = "Input Manual Batch"

        objItem = Nothing
    End Sub

    Private Sub subGoodReceiptAddObject(ByRef pForm As SAPbouiCOM.Form, ByRef pFirstLoad As Boolean)
        Dim objItem As SAPbouiCOM.Item

        'add button
        objItem = pForm.Items.Add("btnCopyGI", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 150
        objItem.Width = 90
        objItem.Specific.Caption = "Copy From"

        objItem = pForm.Items.Add("btnToGI", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 255
        objItem.Width = 90
        objItem.Specific.Caption = "Generate GI"

        objItem = pForm.Items.Add("txtBatch", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 310
        objItem.Width = 150

        objItem = pForm.Items.Add("btnCalc", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        objItem.Top = 318
        objItem.Height = 19
        objItem.Left = 380
        objItem.Width = 100
        objItem.Specific.Caption = "Calculate Quantity"

        'objItem = pForm.Items.Add("txtFishQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        'objItem.Top = 23
        'objItem.Height = 15
        'objItem.Left = 240
        'objItem.Width = 62

        'objItem = pForm.Items.Add("FishQty", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'objItem.Specific.caption = "Fish Qty"
        'objItem.Top = 23
        'objItem.Height = 15
        'objItem.Left = 182
        'objItem.Width = 60
        pForm.Items.Item("btnToGI").Visible = False
        pForm.Items.Item("btnCopyGI").Visible = False
        pForm.Items.Item("txtBatch").Visible = False
        objItem = Nothing
    End Sub

    Public Sub subScrPaint(ByVal pFile As String, ByVal pFormId As String, _
                        ByRef pCounter As Integer, ByVal pForm As SAPbouiCOM.Form)

        Dim strScrPaintLoc As String
        Dim oXML As MSXML2.DOMDocument = Nothing

        strScrPaintLoc = System.Windows.Forms.Application.StartupPath & "\" & pFile

        oXML = New MSXML2.DOMDocument

        oXML.load(strScrPaintLoc)
        oXML.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
            oXML.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & pCounter

        pCounter = pCounter + 1

        objApplication.LoadBatchActions(oXML.xml)

        pForm = objApplication.Forms.GetForm(pFormId, 0)

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXML)
        oXML = Nothing
    End Sub

    Private Sub subCreateTable()
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD = Nothing
        Dim strErrMsg As String
        Dim intNoArray, intLoop As Integer
        Dim arrTableCodeNew(19) As String
        Dim arrTableNameNew(19) As String
        Dim arrTableTypeNew(19) As SAPbobsCOM.BoUTBTableType
        Dim arrTableCode(22) As String

        Try
            oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            '===============================
            'CREATE TABLE
            '===============================
            arrTableCodeNew(0) = ("MIS_WHAUTH")
            arrTableNameNew(0) = ("Warehouse access authorization")
            arrTableTypeNew(0) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(1) = ("MIS_WHDOCS")
            arrTableNameNew(1) = ("Warehouse doc series")
            arrTableTypeNew(1) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(2) = ("MIS_FARMMS")
            arrTableNameNew(2) = ("Farm Master")
            arrTableTypeNew(2) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(3) = ("MIS_NETMS")
            arrTableNameNew(3) = ("Net Master")
            arrTableTypeNew(3) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(4) = ("MIS_PRJMSTR")
            arrTableNameNew(4) = ("Project Master")
            arrTableTypeNew(4) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(5) = ("ReasonCode")
            arrTableNameNew(5) = ("Reason Code")
            arrTableTypeNew(5) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(6) = ("MIS_FSURVR")
            arrTableNameNew(6) = ("Fish survival rate")
            arrTableTypeNew(6) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(7) = ("MIS_FINGB")
            arrTableNameNew(7) = ("Fingerling batch code")
            arrTableTypeNew(7) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(8) = ("MIS_BRAND")
            arrTableNameNew(8) = ("BRAND")
            arrTableTypeNew(8) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(9) = ("MIS_SPEC")
            arrTableNameNew(9) = ("Species")
            arrTableTypeNew(9) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(10) = ("MIS_SKIN")
            arrTableNameNew(10) = ("Skinning")
            arrTableTypeNew(10) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(11) = ("MIS_CUT")
            arrTableNameNew(11) = ("Cutting")
            arrTableTypeNew(11) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(12) = ("MIS_TREAT")
            arrTableNameNew(12) = ("Treatment")
            arrTableTypeNew(12) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(13) = ("MIS_COND")
            arrTableNameNew(13) = ("Condition")
            arrTableTypeNew(13) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(14) = ("MIS_BAG")
            arrTableNameNew(14) = ("Bagging")
            arrTableTypeNew(14) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(15) = ("MIS_GLAZ")
            arrTableNameNew(15) = ("Glazing")
            arrTableTypeNew(15) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(16) = ("MIS_GRADE")
            arrTableNameNew(16) = ("Grade")
            arrTableTypeNew(16) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(17) = ("MIS_LOWLM")
            arrTableNameNew(17) = ("Size Lower Limit")
            arrTableTypeNew(17) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(18) = ("MIS_UPLMT")
            arrTableNameNew(18) = ("Size Upper Limit")
            arrTableTypeNew(18) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            arrTableCodeNew(19) = ("MIS_SIZE")
            arrTableNameNew(19) = ("Size Unit")
            arrTableTypeNew(19) = SAPbobsCOM.BoUTBTableType.bott_NoObject

            intNoArray = 19

            'first run start from 0, if revision just start from last new table
            For intLoop = 0 To intNoArray
                If Not oUserTablesMD.GetByKey(arrTableCodeNew(intLoop)) Then
                    subCreateTableSBO(arrTableCodeNew(intLoop), arrTableNameNew(intLoop), arrTableTypeNew(intLoop))
                End If
            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing

            '=====================================
            'Create Fields
            '=====================================
            arrTableCode(0) = ("MIS_WHAUTH")
            arrTableCode(1) = ("MIS_WHDOCS")
            arrTableCode(2) = ("MIS_FARMMS")
            arrTableCode(3) = ("MIS_NETMS")
            arrTableCode(4) = ("MIS_PRJMSTR")
            arrTableCode(5) = ("OIGE")
            arrTableCode(6) = ("IGE1")
            arrTableCode(7) = ("OITM")
            arrTableCode(8) = ("ReasonCode")
            arrTableCode(9) = ("MIS_FSURVR")
            arrTableCode(10) = ("MIS_FINGB")
            arrTableCode(11) = ("MIS_BRAND")
            arrTableCode(12) = ("MIS_SPEC")
            arrTableCode(13) = ("MIS_SKIN")
            arrTableCode(14) = ("MIS_CUT")
            arrTableCode(15) = ("MIS_TREAT")
            arrTableCode(16) = ("MIS_COND")
            arrTableCode(17) = ("MIS_BAG")
            arrTableCode(18) = ("MIS_GLAZ")
            arrTableCode(19) = ("MIS_GRADE")
            arrTableCode(20) = ("MIS_LOWLM")
            arrTableCode(21) = ("MIS_UPLMT")
            arrTableCode(22) = ("MIS_SIZE")



            intLoop = -1

            'Fields for table MIS_WHAUTH - Warehouse access authorization
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISUSRID", "User id", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISUSRID", "User id", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISWHSCD", "Warehouse code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status  (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table MIS_WHDOCS - Warehouse doc series
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISWHSCD", "Warehouse code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISWHSCD", "Warehouse code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTRXTP", "Transaction type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 50, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISYEARC", "Fiscal year", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 4, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDOCSE", "Document Series", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status  (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If


            'Fields for table MIS_FARMMS - Farm Master
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISFARMC", "Farm code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 3, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMC", "Farm code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 3, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMN", "Farm name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 30, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMA", "Farm address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, SAPbobsCOM.BoYesNoEnum.tNO, 512, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLOCCD", "Location Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMG", "GPS Location", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 128, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status  (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If


            'Fields for table MIS_NETMS - Net Master
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISNETCD", "Net code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETCD", "Net code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMC", "Farm code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 3, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETTP", "Net Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREGCD", "Region", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETLO", "Net location", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 128, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETCA", "Net capacity (m3)", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 11, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETPU", "Net Purposes", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Net Status (O/C)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table MIS_PRJMSTR - Project Master
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            'If fctCheckFieldSBO(arrTableCode(intLoop), "MISPROSR", "Project Stocking remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, , "", True, False) = "" Then
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISPROID", "Project code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 17, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISPROID", "Project code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 17, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETID", "Net code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSIGND", "Stocking date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSCIES", "Species", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)

                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISESTSF", "Estimasi no of fish", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISHARVP", "Day of Culture in days", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 4, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISAGETR", "Age of transfer fish from other Net  *1)", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 4, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISESTHD", "Estimated Harversting date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISESTLF", "Survival Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISESTHQ", "Estimated Harversting qty", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNFDIE", "Cumulative Mortality", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)


                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETPU", "Net purposes", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGENCD", "Fingerling Batch code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 50, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISPROSR", "Project Stocking remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)

                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISHARVD", "Actual Harvesting date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISPROHR", "Project Harvesting remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                '                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGENET", "Genetic Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 50, strErrMsg)

                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISINIFQ", "Initial Fish Qty in Kg", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFEEDQ", "Feed consumption in Kg", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISINIFC", "Initial Fish Cost(Transfer from other net)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFEEDC", "Feed consumption Cost", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)

                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTPCST", "Total Harveting Cost Estimated", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTPGRC", "Total Harvesting good receipt cost - actual", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTPGRQ", "Total Harvesting good receipt qty - actual", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISPROCS", "Project Calculation Flag(0/1)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLCOST", "Labour cost", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDOVCS", "Direct Overhead Cost", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISIOVCS", "Indirect Overhead Cost", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISHARVQ", "Actual Harversting qty", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETST", "Net status  (Open / Closed)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table OIGE - Good Issue
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISTRXTP", "Transaction(Type)", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 1, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTRXTP", "Transaction(Type)", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTRXWH", "Transaction(warehouse)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 100, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESTW", "Destination(warehouse)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 100, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDRVNM", "Driver(name)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 25, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISASDRV", "Ass driver name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 25, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLICNO", "License(no)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 25, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSPVID", "Supervisor(id)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 25, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRITNO", "Rit(number)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREASC", "Reason(code)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 15, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREFFD", "Refer document (GR / GI No.)", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 11, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDELTM", "Delivery(time)", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISARRTM", "Arrival(time)", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLASTB", "Last Harversting batch ?  (Y/N)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISPRNO", "PR No.", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 10, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREQBY", "Request By User", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 254, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTROUT", "Transfer Out Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 10, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTOWHS", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 6, strErrMsg)

            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table IGE1 - Good Issue Row
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISBOXNO", "Box(No)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISBOXNO", "Box(No)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFISHQ", "Fish(Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFRESQ", "Fresh(Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISMORTQ", "Mortality(Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISUNDSQ", "Undersize(Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDEFOQ", "Deformed(Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGISQK", "Good Issue Qty in kg", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGISQP", "Good Issue Qty in pcs", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISVARQK", "Variance Qty in kg", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISVARQP", "Variance Qty in pcs", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISYIELD", "Production(Yield)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNFPRO", "Total convertion fish (system)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNETID", "Net Id(Harversting)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 8, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSEALC", "Seal condition(Good / Open)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISNOSEG", "Seal No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 10, strErrMsg)

            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table OITM - Item Master
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISBRAND", "(Brand)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISBRAND", "(Brand)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSPECS", "(Species)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSKINN", "(Skinning)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCUTTI", "(Cutting)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSIZET", "Size tag (size lower limit + upper limmit + size unit)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISTREAT", "(Treatment)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCONDI", "(Condition(freezing / fresh))", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISBAGGI", "(Bagging)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGLAZE", "(Glazing)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISGRADE", "(Grade)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISYIELD", "Production std yield", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSLLIM", "Size lower limit", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSULIM", "Size upper limit", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSUNIT", "Size(unit)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 2, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFGRDF", "First grade Flag only (unpack finish goods)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table ReasonCode
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISREASC", "Good Issue reason code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREASC", "Good Issue reason code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 15, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISREASD", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 35, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status  (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If


            'Fields for table Fish Survival Rate
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISFARMC", "Farm code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 3, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFARMC", "Farm code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 3, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFSPEC", "Species", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSURVR", "Survival rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, , strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status  (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for table Fingerling batch code
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISFINGB", "Fingerling Batch ", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 50, "", True, False) = "" Then
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISFINGB", "Fingerling Batch ", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 50, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISSTRN", "Strain", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 100, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISACRO", "Acronym", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 10, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESC", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 254, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Brand
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Species
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Skinning
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Cutting
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Treatment
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Condition
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Bagging
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Glazing
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Grade
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for Size lower limit
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

            'Fields for size upper limit
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If


            'Fields for Size unit
            strErrMsg = ""
            intLoop = intLoop + 1

            'cek latest field, if exists assume all fields already create
            If fctCheckFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, "", True, False) = "" Then
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISCODE", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, 2, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISDESCR", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 9, strErrMsg)
                strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISRECST", "Record Status (A/D)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 1, strErrMsg)
                'strErrMsg = fctCreateFieldSBO(arrTableCode(intLoop), "MISLENGH", "Length", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, 5, strErrMsg)
            End If

            If strErrMsg <> "" Then
                MsgBox("Fail adding fields: " & strErrMsg & " into table " & arrTableCode(intLoop), vbExclamation, "SAP BO")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SAP BO")
        Finally
            objApplication.StatusBar.SetText("Set Table Successfully.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub

    Private Sub subCreateTableSBO(ByVal pName As String, ByVal pDesc As String, ByVal pType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD = Nothing
        Dim lngRetCodeTbl As Long

        oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

        oUserTablesMD.TableName = pName
        oUserTablesMD.TableDescription = Left(pDesc, 30)
        oUserTablesMD.TableType = pType
        lngRetCodeTbl = oUserTablesMD.Add

        If lngRetCodeTbl <> 0 And (lngRetCodeTbl <> -2035) Then
            MsgBox("Fail adding table " & pName, vbExclamation, "SAP BO")
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
        oUserTablesMD = Nothing
    End Sub

    Private Function fctCheckFieldSBO(ByVal pTable As String, ByVal pField As String, ByVal pDesc As String, _
                                    ByVal pType As SAPbobsCOM.BoFieldTypes, ByVal pSubType As SAPbobsCOM.BoFldSubTypes, ByVal pMandatory As SAPbobsCOM.BoYesNoEnum, _
                                    Optional ByVal pEditSize As Integer = 0, Optional ByVal pErrMsg As String = "", Optional ByVal pExisting As Boolean = False, _
                                    Optional ByVal pSAPTable As Boolean = False) As String

        Dim oRecSet As SAPbobsCOM.Recordset = Nothing

        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If pSAPTable Then
            oRecSet.DoQuery("Select TABLEID FROM CUFD WHERE TABLEID = '" & pTable & "' AND ALIASID = '" & pField & "'")
        Else
            oRecSet.DoQuery("Select TABLEID FROM CUFD WHERE TABLEID = '@" & pTable & "' AND ALIASID = '" & pField & "'")
        End If

        If oRecSet.RecordCount > 0 Then
            fctCheckFieldSBO = IIf(pErrMsg <> "", pErrMsg + ", " + pField, pField)
        Else
            fctCheckFieldSBO = pErrMsg
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        oRecSet = Nothing
    End Function

    Private Function fctCreateFieldSBO(ByVal pTable As String, ByVal pField As String, ByVal pDesc As String, _
                                    ByVal pType As SAPbobsCOM.BoFieldTypes, ByVal pSubType As SAPbobsCOM.BoFldSubTypes, ByVal pMandatory As SAPbobsCOM.BoYesNoEnum, _
                                    Optional ByVal pEditSize As Integer = 0, Optional ByVal pErrMsg As String = "", Optional ByVal pExisting As Boolean = False) As String

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        Dim lngRetCodeFld As Long

        fctCreateFieldSBO = ""

        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = pTable
        oUserFieldsMD.Name = pField
        oUserFieldsMD.Description = Left(pDesc, 30)
        oUserFieldsMD.Type = pType
        oUserFieldsMD.SubType = pSubType
        oUserFieldsMD.Mandatory = pMandatory

        If pMandatory Then
            Select Case pType
                Case SAPbobsCOM.BoFieldTypes.db_Date
                    'reset mandatory to false, because there is difference
                    oUserFieldsMD.Mandatory = False
                    'oUserFieldsMD.DefaultValue = "getdate()"
                Case SAPbobsCOM.BoFieldTypes.db_Alpha
                    'reset mandatory to false, because there is difference
                    oUserFieldsMD.Mandatory = False
                    'oUserFieldsMD.DefaultValue = "('')"
                Case SAPbobsCOM.BoFieldTypes.db_Numeric
                    oUserFieldsMD.DefaultValue = "0"
                Case SAPbobsCOM.BoFieldTypes.db_Float
                    oUserFieldsMD.DefaultValue = "0"
            End Select
        End If

        'Type : 0 = db_Alpha, 2 = db_Numeric
        If pType = SAPbobsCOM.BoFieldTypes.db_Alpha Or pType = SAPbobsCOM.BoFieldTypes.db_Numeric Then
            oUserFieldsMD.EditSize = pEditSize
        End If

        lngRetCodeFld = oUserFieldsMD.Add
        If lngRetCodeFld <> 0 And (lngRetCodeFld <> -2035 And lngRetCodeFld <> -1) Then
            fctCreateFieldSBO = IIf(pErrMsg <> "", pErrMsg + ", " + pField, pField)
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
        oUserFieldsMD = Nothing
    End Function

    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        Dim objForm As SAPbouiCOM.Form
        Dim intForm As Integer

        objForm = objApplication.Forms.ActiveForm

        If pVal.BeforeAction Then

            Select Case pVal.MenuUID
                Case ProjectMaster_MenuId
                    Try
                        If fctFormExist(ProjectMaster_FormId, intForm) Then
                            objApplication.Forms.Item(intForm).Select()
                        Else
                            subProjectMasterScrPaint()
                            SubProjectMasterModeAdd()
                        End If
                    Catch ex As Exception
                        Beep()
                    End Try

                Case ProjectHarvest_MenuId
                    Try
                        If fctFormExist(ProjectHarvest_FormId, intForm) Then
                            objApplication.Forms.Item(intForm).Select()
                        Else
                            subProjectHarvestScrPaint()
                            SubProjectHarvestModeFind()
                        End If
                    Catch ex As Exception
                        Beep()
                    End Try


                Case TBar_Remove
                    If objForm.Type = ProjectMaster_FormId Then
                        SubRemove(objForm)
                    End If

                Case TBar_Update
                    If objForm.Type = ProjectMaster_FormId Then
                        SubUpdateStatus(objForm)
                    End If

            End Select

            If objForm.Type = ProjectMaster_FormId Then
                Select Case pVal.MenuUID
                    Case TBar_First
                        SubToolbarAction(TBar_First, objForm)

                    Case TBar_Last
                        SubToolbarAction(TBar_Last, objForm)

                    Case TBar_Prev
                        SubToolbarAction(TBar_Prev, objForm)

                    Case TBar_Next
                        SubToolbarAction(TBar_Next, objForm)

                    Case TBar_Find

                        SubToolbarAction(TBar_Find, objForm)
                        BubbleEvent = False

                    Case TBar_Add
                        SubToolbarAction(TBar_Add, objForm)
                        BubbleEvent = False
                End Select

            ElseIf objForm.Type = ProjectHarvest_FormId Then
                Select Case pVal.MenuUID
                    Case TBar_First
                        SubToolbarAction(TBar_First, objForm)

                    Case TBar_Last
                        SubToolbarAction(TBar_Last, objForm)

                    Case TBar_Prev
                        SubToolbarAction(TBar_Prev, objForm)

                    Case TBar_Next
                        SubToolbarAction(TBar_Next, objForm)

                    Case TBar_Find

                        SubToolbarAction(TBar_Find, objForm)
                        BubbleEvent = False

                    Case TBar_Add
                        SubToolbarAction(TBar_Add, objForm)
                        BubbleEvent = False
                End Select
            End If

        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm)

    End Sub

    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        Dim oCreationPackage1 As SAPbouiCOM.MenuCreationParams
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenu As SAPbouiCOM.Menus
        Dim DeleteDocEntry As String

        If Not (objFormProjectMaster Is Nothing) Then
            If eventInfo.FormUID = objFormProjectMaster.UniqueID And eventInfo.EventType = SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK Then
                If UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption) <> "UPDATE" And UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption) <> "OK" Then GoTo Setnothing
                If eventInfo.ItemUID = "MISPROID" Then
                    If eventInfo.BeforeAction Then
                        DeleteDocEntry = objFormProjectMaster.Items.Item("MISPROID").Specific.String

                        oCreationPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = TBar_Remove
                        oCreationPackage.String = TBar_Remove

                        oCreationPackage1 = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage1.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage1.UniqueID = TBar_Update
                        oCreationPackage1.String = TBar_Update

                        oCreationPackage.Enabled = True
                        oCreationPackage1.Enabled = True

                        oMenuItem = objApplication.Menus.Item(Tbar_Data)
                        oMenu = oMenuItem.SubMenus
                        If Not oMenu.Exists(TBar_Remove) Then
                            oMenu.AddEx(oCreationPackage)
                            oMenu.AddEx(oCreationPackage1)
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage1)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenu)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem)

                    Else
                        oMenuItem = objApplication.Menus.Item(Tbar_Data)
                        oMenu = oMenuItem.SubMenus
                        oMenu.RemoveEx(TBar_Remove)


                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenu)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem)
                    End If
                End If
            End If
        End If

Setnothing:
        oMenu = Nothing
        oMenuItem = Nothing

    End Sub

    Private Sub SubUpdateStatus(ByVal pForm As SAPbouiCOM.Form)
        Dim StrSql As String
        Dim Code As Integer
        Dim NetId As String
        Dim ProjectId As String
        Dim DocDate As Date
        Dim ObjRecSet As SAPbobsCOM.Recordset = Nothing


        ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Code = pForm.Items.Item("Code").Specific.string
        NetId = pForm.Items.Item("MISNETID").Specific.string
        ProjectId = pForm.Items.Item("MISPROID").Specific.string
        DocDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, pForm.Items.Item("MISSIGND").Specific.string, 1))


        StrSql = "Update [@MIS_PRJMSTR] SET [U_MISNETST] = 'H' Where DocEntry = " & Code & ""
        ObjRecSet.DoQuery(StrSql)

        If ObjRecSet.RecordCount = 0 Then
            pForm.Items.Item("MISNETST").Specific.string = "H"
            pForm.Items.Item("btnOK").Specific.caption = "OK"
            objApplication.StatusBar.SetText("Update Status SuccessFull ~18.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Else
            objApplication.StatusBar.SetText("Update Status Not SuccessFull ~18.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)
        ObjRecSet = Nothing
    End Sub

    Private Sub SubRemove(ByVal pForm As SAPbouiCOM.Form)
        Dim StrSql As String
        Dim Code As Integer
        Dim NetId As String
        Dim ProjectId As String
        Dim DocDate As Date
        Dim ObjRecSet As SAPbobsCOM.Recordset = Nothing

        ObjRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Code = pForm.Items.Item("Code").Specific.string
        NetId = pForm.Items.Item("MISNETID").Specific.string
        ProjectId = pForm.Items.Item("MISPROID").Specific.string
        DocDate = CDate(ClsGlobal.fctFormatDateSave(oCompany, pForm.Items.Item("MISSIGND").Specific.string, 1))

        'StrSql = "SELECT * FROM IGE1 WHERE U_MISNETID = '" & NetId & "' AND DocDate = '" & DocDate & "' "
        StrSql = "SELECT * FROM IGE1 WHERE U_MISPROID = '" & ProjectId & "' "
        ObjRecSet.DoQuery(StrSql)

        If ObjRecSet.RecordCount > 0 Then
            objApplication.StatusBar.SetText("Cannot Delete Because there were already good issue transaction ~18.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

        StrSql = "Update [@MIS_PRJMSTR] SET [U_MISNETST] = 'D' Where DocEntry = " & Code & ""
        ObjRecSet.DoQuery(StrSql)

        If ObjRecSet.RecordCount = 0 Then
            pForm.Items.Item("MISNETST").Specific.string = "D"
            pForm.Items.Item("btnOK").Specific.caption = "OK"
            objApplication.StatusBar.SetText("Delete Status SuccessFull ~18.0003~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Else
            objApplication.StatusBar.SetText("Update Status Not SuccessFull ~18.0002~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            GoTo setnothing
        End If

setnothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRecSet)
        ObjRecSet = Nothing
    End Sub

    Private Sub SubToolbarAction(ByVal pToolBar As String, ByVal objForm As SAPbouiCOM.Form)

        Dim StrFormId As String

        objForm = objApplication.Forms.ActiveForm
        StrFormId = objForm.TypeEx

        If StrFormId = ProjectMaster_FormId Then
            Select Case pToolBar

                Case TBar_First
                    SubProjectMasterToolbar()
                    SubModeDisplayData(TBar_First, StrFormId)

                Case TBar_Last
                    SubProjectMasterToolbar()
                    SubModeDisplayData(TBar_Last, StrFormId)

                Case TBar_Prev
                    SubProjectMasterToolbar()
                    SubModeDisplayData(TBar_Prev, StrFormId)

                Case TBar_Next
                    SubProjectMasterToolbar()
                    SubModeDisplayData(TBar_Next, StrFormId)

                Case TBar_Find
                    Select Case StrFormId
                        Case ProjectMaster_FormId
                            SubProjectMasterModeFind()
                    End Select

                Case TBar_Add
                    Select Case StrFormId
                        Case ProjectMaster_FormId
                            SubProjectMasterModeAdd()
                    End Select
            End Select

        ElseIf objForm.Type = ProjectHarvest_FormId Then
            Select Case pToolBar

                Case TBar_First
                    SubProjectHarvestToolbar()
                    SubModeDisplayData(TBar_First, StrFormId)

                Case TBar_Last
                    SubProjectHarvestToolbar()
                    SubModeDisplayData(TBar_Last, StrFormId)

                Case TBar_Prev
                    SubProjectHarvestToolbar()
                    SubModeDisplayData(TBar_Prev, StrFormId)

                Case TBar_Next
                    SubProjectHarvestToolbar()
                    SubModeDisplayData(TBar_Next, StrFormId)

                Case TBar_Find
                    Select Case StrFormId
                        Case ProjectHarvest_FormId
                            SubProjectHarvestModeFind()
                    End Select

                Case TBar_Add
                    Select Case StrFormId
                        Case ProjectHarvest_FormId
                            SubProjectHarvestModeAdd()
                    End Select
            End Select
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(objForm)


    End Sub

    Private Sub SubProjectHarvestToolbar()
        'oFormProjectHarvest.Items.Item("btnOK").Specific.Caption = "FIND"

        SubSetToolbar(oFormProjectHarvest, True, True, True, False, True, True, True, True, _
                                False, False, False, False, True, True, True)
    End Sub

    Private Sub SubProjectMasterToolbar()
        objFormProjectMaster.Items.Item("btnOK").Specific.Caption = "UPDATE"

        SubSetToolbar(objFormProjectMaster, True, True, True, True, True, True, True, True, _
                                False, False, False, False, True, True, True)

    End Sub

    Private Sub SubProjectHarvestModeAdd()
        oFormProjectHarvest.Freeze(True)
        ' Dim StrCaption As String
        SubFPClearObjValue(oFormProjectHarvest, UCase(oFormProjectHarvest.Items.Item("btnOK").Specific.Caption))
        'SubFPLoadObjValue(objFormProjectMaster)

        oFormProjectHarvest.Items.Item("MISNETID").Click()
        oFormProjectHarvest.Items.Item("MISSCIES").Enabled = True
        oFormProjectHarvest.Items.Item("MISESTSF").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVP").Enabled = False
        oFormProjectHarvest.Items.Item("MISAGETR").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTHD").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTLF").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTHQ").Enabled = False
        oFormProjectHarvest.Items.Item("MISNFDIE").Enabled = False
        oFormProjectHarvest.Items.Item("MISNETPUCD").Enabled = False
        oFormProjectHarvest.Items.Item("MISGENCD").Enabled = False
        oFormProjectHarvest.Items.Item("MISPROSR").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVD").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVQ").Enabled = False
        oFormProjectHarvest.Items.Item("MISPROHR").Enabled = False

        SubSetToolbar(oFormProjectHarvest, True, False, True, False, True, True, True, True, _
                                        False, False, False, False, True, True, True)

        oFormProjectHarvest.Freeze(False)

        'subObjEnabledFP(objFormProjectMaster, "ADD") 'StrCaption

    End Sub

    Private Sub SubProjectMasterModeAdd()
        objFormProjectMaster.Freeze(True)

        objFormProjectMaster.Items.Item("btnOK").Specific.Caption = "ADD"
        ' Dim StrCaption As String
        SubFPClearObjValue(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption))
        'SubFPLoadObjValue(objFormProjectMaster)

        objFormProjectMaster.Items.Item("MISNETID").Enabled = True
        objFormProjectMaster.Items.Item("MISNETID").Click()
        objFormProjectMaster.Items.Item("MISSIGND").Enabled = True
        objFormProjectMaster.Items.Item("MISSCIES").Enabled = True
        objFormProjectMaster.Items.Item("MISESTSF").Enabled = True
        objFormProjectMaster.Items.Item("MISHARVP").Enabled = False
        objFormProjectMaster.Items.Item("MISAGETR").Enabled = False
        objFormProjectMaster.Items.Item("MISESTHD").Enabled = False
        objFormProjectMaster.Items.Item("MISESTLF").Enabled = False
        objFormProjectMaster.Items.Item("MISESTHQ").Enabled = False
        objFormProjectMaster.Items.Item("MISNFDIE").Enabled = False
        objFormProjectMaster.Items.Item("MISNETPUCD").Enabled = False
        objFormProjectMaster.Items.Item("MISGENCD").Enabled = True
        objFormProjectMaster.Items.Item("MISPROSR").Enabled = True
        objFormProjectMaster.Items.Item("MISHARVD").Enabled = False
        objFormProjectMaster.Items.Item("MISHARVQ").Enabled = False
        objFormProjectMaster.Items.Item("MISPROHR").Enabled = False

        SubSetToolbar(objFormProjectMaster, True, False, True, False, True, True, True, True, _
                                        False, False, False, False, True, True, True)

        objFormProjectMaster.Freeze(False)

        'subObjEnabledFP(objFormProjectMaster, "ADD") 'StrCaption

    End Sub

    Private Sub SubProjectHarvestModeFind()

        oFormProjectHarvest.Items.Item("btnOK").Specific.Caption = "FIND"
        ' Dim StrCaption As String
        SubFPClearObjValue(oFormProjectHarvest, UCase(oFormProjectHarvest.Items.Item("btnOK").Specific.Caption))
        'SubFPLoadObjValue(objFormProjectMaster)

        oFormProjectHarvest.Items.Item("MISNETID").Enabled = True
        oFormProjectHarvest.Items.Item("MISNETID").Click()
        oFormProjectHarvest.Items.Item("MISSIGND").Enabled = True
        oFormProjectHarvest.Items.Item("MISSCIES").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTSF").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVP").Enabled = False
        oFormProjectHarvest.Items.Item("MISAGETR").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTHD").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTLF").Enabled = False
        oFormProjectHarvest.Items.Item("MISESTHQ").Enabled = False
        oFormProjectHarvest.Items.Item("MISNFDIE").Enabled = False
        oFormProjectHarvest.Items.Item("MISNETPUCD").Enabled = False
        oFormProjectHarvest.Items.Item("MISGENCD").Enabled = False
        oFormProjectHarvest.Items.Item("MISPROSR").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVD").Enabled = False
        oFormProjectHarvest.Items.Item("MISHARVQ").Enabled = False
        oFormProjectHarvest.Items.Item("MISPROHR").Enabled = False

        SubSetToolbar(oFormProjectHarvest, True, False, True, False, True, True, True, True, _
                                        False, False, False, False, True, True, True)


    End Sub

    Private Sub SubProjectMasterModeFind()

        objFormProjectMaster.Items.Item("btnOK").Specific.Caption = "FIND"
        ' Dim StrCaption As String
        SubFPClearObjValue(objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption))
        'SubFPLoadObjValue(objFormProjectMaster)

        objFormProjectMaster.Items.Item("MISNETID").Enabled = True
        objFormProjectMaster.Items.Item("MISNETID").Click()
        objFormProjectMaster.Items.Item("MISSIGND").Enabled = True
        objFormProjectMaster.Items.Item("MISSCIES").Enabled = False
        objFormProjectMaster.Items.Item("MISESTSF").Enabled = False
        objFormProjectMaster.Items.Item("MISHARVP").Enabled = False
        objFormProjectMaster.Items.Item("MISAGETR").Enabled = False
        objFormProjectMaster.Items.Item("MISESTHD").Enabled = False
        objFormProjectMaster.Items.Item("MISESTLF").Enabled = False
        objFormProjectMaster.Items.Item("MISESTHQ").Enabled = False
        objFormProjectMaster.Items.Item("MISNFDIE").Enabled = False
        objFormProjectMaster.Items.Item("MISNETPUCD").Enabled = False
        objFormProjectMaster.Items.Item("MISGENCD").Enabled = False
        objFormProjectMaster.Items.Item("MISPROSR").Enabled = False
        objFormProjectMaster.Items.Item("MISHARVD").Enabled = False
        objFormProjectMaster.Items.Item("MISHARVQ").Enabled = False
        objFormProjectMaster.Items.Item("MISPROHR").Enabled = False

        SubSetToolbar(objFormProjectMaster, True, False, False, True, True, True, True, True, _
                                        False, False, False, False, True, True, True)


    End Sub

    Private Sub SubFPClearObjValue(ByVal pForm As SAPbouiCOM.Form, ByVal pMode As String)
        On Error GoTo ErrorHandler

        pForm.Freeze(True)
        pForm.Items.Item("Code").Specific.string = ""
        pForm.Items.Item("MISPROID").Specific.string = ""
        pForm.Items.Item("MISNETID").Specific.string = ""
        pForm.Items.Item("MISSIGND").Specific.string = ""
        pForm.Items.Item("MISSCIES").Specific.string = ""
        pForm.Items.Item("MISESTSF").Specific.string = 0
        pForm.Items.Item("MISHARVP").Specific.string = 0
        pForm.Items.Item("MISAGETR").Specific.string = 0
        pForm.Items.Item("MISESTHD").Specific.string = ""
        pForm.Items.Item("MISESTLF").Specific.string = 100
        pForm.Items.Item("MISESTHQ").Specific.string = 0
        pForm.Items.Item("MISNFDIE").Specific.string = 0
        pForm.Items.Item("MISNETPUCD").Specific.string = ""
        pForm.Items.Item("MISGENCD").Specific.string = ""
        pForm.Items.Item("MISGENET").Specific.string = ""
        pForm.Items.Item("MISPROSR").Specific.string = ""
        pForm.Items.Item("MISHARVD").Specific.string = ""
        pForm.Items.Item("MISHARVQ").Specific.string = 0
        pForm.Items.Item("MISPROHR").Specific.string = ""
        pForm.Items.Item("MISINIFQ").Specific.string = 0
        pForm.Items.Item("MISFEEDQ").Specific.string = 0
        pForm.Items.Item("MISINIFC").Specific.string = 0
        pForm.Items.Item("MISFEEDC").Specific.string = 0
        pForm.Items.Item("MISTPCST").Specific.string = 0
        pForm.Items.Item("MISTPGRC").Specific.string = 0
        pForm.Items.Item("MISTPGRQ").Specific.string = 0
        pForm.Items.Item("MISPROCS").Specific.string = ""
        pForm.Items.Item("MISNETST").Specific.string = "O"
        pForm.Items.Item("MISHATGO").Specific.string = ""
        pForm.Items.Item("MISFCR").Specific.string = 0
        pForm.Items.Item("MISFCE").Specific.string = 0
        pForm.Items.Item("MISTEFQK").Specific.string = 0
        pForm.Freeze(False)

ErrorHandler:
        If Err.Number <> 0 Then
            MsgBox("Fail Clear Object !~90001~", vbExclamation, "SAP BO")
        End If


    End Sub

    Private Sub SubModeDisplayData(ByVal pToolBar As String, ByVal pFormId As String)
        Dim objRecSet As SAPbobsCOM.Recordset = Nothing
        Dim StrSql As String
        Dim StrMsg As String

        objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        'Select Case pFormId
        '    Case ProjectMaster_FormId
        '        StrSql = "Select Top 1 T0.* " & _
        '                " From [@MIS_PRJMSTR] T0 "
        'End Select

        If (pToolBar = TBar_Prev) Or (pToolBar = TBar_Next) Then
            Select Case pToolBar
                Case TBar_Prev
                    Select Case pFormId
                        Case ProjectMaster_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0 Where T0.DocEntry < " + IIf(Trim(objFormProjectMaster.Items.Item("Code").Specific.string) = "", "0", objFormProjectMaster.Items.Item("Code").Specific.string) + _
                                        " order by T0.DocEntry desc"

                            objRecSet.DoQuery(StrSql)

                        Case ProjectHarvest_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0 Where T0.DocEntry < " + IIf(Trim(oFormProjectHarvest.Items.Item("Code").Specific.string) = "", "0", oFormProjectHarvest.Items.Item("Code").Specific.string) + _
                                        " order by T0.DocEntry desc"

                            objRecSet.DoQuery(StrSql)
                    End Select


                    StrMsg = "Last Record."
                    pToolBar = TBar_Last

                Case TBar_Next
                    Select Case pFormId
                        Case ProjectMaster_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0 Where T0.DocEntry > " + IIf(Trim(objFormProjectMaster.Items.Item("Code").Specific.string) = "", "0", objFormProjectMaster.Items.Item("Code").Specific.string) + _
                                        " order by T0.DocEntry asc"

                            objRecSet.DoQuery(StrSql)

                        Case ProjectHarvest_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0 Where T0.DocEntry > " + IIf(Trim(oFormProjectHarvest.Items.Item("Code").Specific.string) = "", "0", oFormProjectHarvest.Items.Item("Code").Specific.string) + _
                                        " order by T0.DocEntry asc"

                            objRecSet.DoQuery(StrSql)
                    End Select

                    StrMsg = "First Record."
                    pToolBar = TBar_First

            End Select

            'objRecSet.DoQuery(StrSql)

            If objRecSet.RecordCount > 0 Then
                StrMsg = ""
                pToolBar = ""
                'Else
                'objApplication.StatusBar.SetText("No Data", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If


        If (pToolBar = TBar_First Or pToolBar = TBar_Last) Then

            Select Case pToolBar
                Case TBar_First
                    Select Case pFormId
                        Case ProjectMaster_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0  order by T0.DocEntry asc"

                            objRecSet.DoQuery(StrSql)

                        Case ProjectHarvest_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0  order by T0.DocEntry asc"

                            objRecSet.DoQuery(StrSql)
                    End Select

                    StrMsg = "First Record."
                    pToolBar = TBar_First

                Case TBar_Last
                    Select Case pFormId
                        Case ProjectMaster_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                        " From [@MIS_PRJMSTR] T0  order by T0.DocEntry desc"

                            objRecSet.DoQuery(StrSql)

                        Case ProjectHarvest_FormId
                            StrSql = "Select Top 1 T0.* " & _
                                    " From [@MIS_PRJMSTR] T0  order by T0.DocEntry desc"

                            objRecSet.DoQuery(StrSql)
                    End Select


                    StrMsg = "Last Record."
                    pToolBar = TBar_Last

            End Select

        End If

        If objRecSet.RecordCount > 0 Then
            Select Case pFormId
                Case ProjectMaster_FormId
                    SubProjectMasterDisplayData("Master", objRecSet, objFormProjectMaster, UCase(objFormProjectMaster.Items.Item("btnOK").Specific.Caption))
                Case ProjectHarvest_FormId
                    SubProjectMasterDisplayData("Harvest", objRecSet, oFormProjectHarvest, UCase(oFormProjectHarvest.Items.Item("btnOK").Specific.Caption))
            End Select
        Else
            objApplication.StatusBar.SetText("No Data ~16.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

        objRecSet = Nothing
    End Sub

    Private Sub SubProjectMasterDisplayData(ByVal Project As String, ByVal pRecSet As SAPbobsCOM.Recordset, ByVal pForm As SAPbouiCOM.Form, ByVal pMode As String)
        Dim TanggalStock As String
        Dim TanggalHarvest As String
        Dim TanggalActual As String
        TanggalStock = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISSIGND").Value, oCompany)
        TanggalHarvest = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISESTHD").Value, oCompany)
        TanggalActual = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISHARVD").Value, oCompany)

        pForm.Freeze(True)

        pForm.Items.Item("Code").Specific.string = pRecSet.Fields.Item("DocEntry").Value
        pForm.Items.Item("MISPROID").Specific.string = pRecSet.Fields.Item("U_MISPROID").Value
        pForm.Items.Item("MISNETID").Specific.string = pRecSet.Fields.Item("U_MISNETID").Value

        If TanggalStock = "30/12/2099" Or TanggalStock = "30.12.2099" Or TanggalStock = "30-12-2099" Then
            pForm.Items.Item("MISSIGND").Specific.string = ""
        Else
            pForm.Items.Item("MISSIGND").Specific.string = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISSIGND").Value, oCompany)
        End If
        pForm.Items.Item("MISSCIES").Specific.string = pRecSet.Fields.Item("U_MISSCIES").Value
        pForm.Items.Item("MISESTSF").Specific.string = pRecSet.Fields.Item("U_MISESTSF").Value
        pForm.Items.Item("MISHARVP").Specific.string = pRecSet.Fields.Item("U_MISHARVP").Value

        If pRecSet.Fields.Item("U_MISNETST").Value = "O" Then
            Dim Datenow As Date = Date.Now
            Dim DateStock As Date = pRecSet.Fields.Item("U_MISSIGND").Value

            pForm.Items.Item("MISAGETR").Specific.string = DateDiff(DateInterval.Day, DateStock, Datenow)
        Else
            pForm.Items.Item("MISAGETR").Specific.string = pRecSet.Fields.Item("U_MISAGETR").Value
        End If

        If TanggalHarvest = "30/12/2099" Or TanggalHarvest = "30.12.2099" Or TanggalHarvest = "30-12-2099" Then
            pForm.Items.Item("MISESTHD").Specific.string = ""
        Else
            pForm.Items.Item("MISESTHD").Specific.string = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISESTHD").Value, oCompany)
        End If
        pForm.Items.Item("MISESTLF").Specific.string = pRecSet.Fields.Item("U_MISESTLF").Value
        'pForm.Items.Item("MISGENET").Specific.string = pRecSet.Fields.Item("U_MISGENET").Value
        pForm.Items.Item("MISESTHQ").Specific.string = pRecSet.Fields.Item("U_MISESTHQ").Value
        pForm.Items.Item("MISNFDIE").Specific.string = pRecSet.Fields.Item("U_MISNFDIE").Value
        pForm.Items.Item("MISNETPUCD").Specific.string = pRecSet.Fields.Item("U_MISNETPU").Value
        pForm.Items.Item("MISGENCD").Specific.string = pRecSet.Fields.Item("U_MISGENCD").Value
        pForm.Items.Item("MISPROSR").Specific.string = pRecSet.Fields.Item("U_MISPROSR").Value
        If TanggalActual = "30/12/2099" Or TanggalActual = "30-12-2099" Or TanggalActual = "30.12.2099" Then
            pForm.Items.Item("MISHARVD").Specific.string = ""
        Else
            pForm.Items.Item("MISHARVD").Specific.string = ClsGlobal.fctFormatDate(pRecSet.Fields.Item("U_MISHARVD").Value, oCompany)
        End If
        pForm.Items.Item("MISHARVQ").Specific.string = pRecSet.Fields.Item("U_MISHARVQ").Value
        pForm.Items.Item("MISPROHR").Specific.string = pRecSet.Fields.Item("U_MISPROHR").Value
        pForm.Items.Item("MISINIFQ").Specific.string = pRecSet.Fields.Item("U_MISINIFQ").Value
        pForm.Items.Item("MISFEEDQ").Specific.string = pRecSet.Fields.Item("U_MISFEEDQ").Value
        pForm.Items.Item("MISFCR").Specific.string = pRecSet.Fields.Item("U_MISFCR").Value
        pForm.Items.Item("MISFCE").Specific.string = pRecSet.Fields.Item("U_MISFCE").Value
        pForm.Items.Item("MISTEFQK").Specific.string = pRecSet.Fields.Item("U_MISTFQKG").Value
        pForm.Items.Item("MISINIFC").Specific.string = pRecSet.Fields.Item("U_MISINIFC").Value
        pForm.Items.Item("MISFEEDC").Specific.string = pRecSet.Fields.Item("U_MISFEEDC").Value
        pForm.Items.Item("MISTPCST").Specific.string = pRecSet.Fields.Item("U_MISTPCST").Value
        pForm.Items.Item("MISTPGRC").Specific.string = pRecSet.Fields.Item("U_MISTPGRC").Value
        pForm.Items.Item("MISTPGRQ").Specific.string = pRecSet.Fields.Item("U_MISTPGRQ").Value
        pForm.Items.Item("MISPROCS").Specific.string = pRecSet.Fields.Item("U_MISPROCS").Value
        pForm.Items.Item("MISNETST").Specific.string = pRecSet.Fields.Item("U_MISNETST").Value
        pForm.Items.Item("MISHATGO").Specific.string = pRecSet.Fields.Item("U_MISHATGRO").Value

        If Project = "Master" Then
            If pForm.Items.Item("btnOK").Specific.Caption = "UPDATE" Then
                If pForm.Items.Item("MISNETST").Specific.string = "O" Then
                    pForm.Items.Item("MISNETID").Click()
                    pForm.Items.Item("MISSIGND").Enabled = False
                    pForm.Items.Item("MISSCIES").Enabled = False
                    pForm.Items.Item("MISESTSF").Enabled = False
                    pForm.Items.Item("MISHARVP").Enabled = True
                    pForm.Items.Item("MISAGETR").Enabled = True
                    pForm.Items.Item("MISESTHD").Enabled = False
                    pForm.Items.Item("MISESTLF").Enabled = False
                    pForm.Items.Item("MISESTHQ").Enabled = False
                    pForm.Items.Item("MISNFDIE").Enabled = False
                    pForm.Items.Item("MISNETPUCD").Enabled = False
                    pForm.Items.Item("MISGENCD").Enabled = True
                    pForm.Items.Item("MISPROSR").Enabled = False
                    pForm.Items.Item("MISHARVD").Enabled = False
                    pForm.Items.Item("MISHARVQ").Enabled = False
                    pForm.Items.Item("MISPROHR").Enabled = False
                ElseIf pForm.Items.Item("MISNETST").Specific.string <> "O" Then
                    pForm.Items.Item("btnOK").Specific.Caption = "OK"
                    pForm.Items.Item("MISNETID").Click()
                    pForm.Items.Item("MISSIGND").Enabled = False
                    pForm.Items.Item("MISSCIES").Enabled = False
                    pForm.Items.Item("MISESTSF").Enabled = False
                    pForm.Items.Item("MISHARVP").Enabled = False
                    pForm.Items.Item("MISAGETR").Enabled = False
                    pForm.Items.Item("MISESTHD").Enabled = False
                    pForm.Items.Item("MISESTLF").Enabled = False
                    pForm.Items.Item("MISESTHQ").Enabled = False
                    pForm.Items.Item("MISNFDIE").Enabled = False
                    pForm.Items.Item("MISNETPUCD").Enabled = False
                    pForm.Items.Item("MISGENCD").Enabled = False
                    pForm.Items.Item("MISPROSR").Enabled = False
                    pForm.Items.Item("MISHARVD").Enabled = False
                    pForm.Items.Item("MISHARVQ").Enabled = False
                    pForm.Items.Item("MISPROHR").Enabled = False
                Else
                    pForm.Items.Item("MISNETID").Click()
                    pForm.Items.Item("MISSCIES").Enabled = False
                    pForm.Items.Item("MISESTSF").Enabled = True
                    pForm.Items.Item("MISHARVP").Enabled = True
                    pForm.Items.Item("MISAGETR").Enabled = True
                    pForm.Items.Item("MISESTHD").Enabled = False
                    pForm.Items.Item("MISESTLF").Enabled = False
                    pForm.Items.Item("MISESTHQ").Enabled = False
                    pForm.Items.Item("MISNFDIE").Enabled = False
                    pForm.Items.Item("MISNETPUCD").Enabled = False
                    pForm.Items.Item("MISGENCD").Enabled = True
                    pForm.Items.Item("MISPROSR").Enabled = False
                    pForm.Items.Item("MISHARVD").Enabled = False
                    pForm.Items.Item("MISHARVQ").Enabled = False
                    pForm.Items.Item("MISPROHR").Enabled = False
                    If pForm.Items.Item("btnOK").Specific.Caption = "FIND" Then
                        pForm.Items.Item("btnOK").Specific.Caption = "UPDATE"
                    End If
                End If
            ElseIf pForm.Items.Item("btnOK").Specific.Caption = "FIND" Then
                pForm.Items.Item("MISNETID").Click()
                pForm.Items.Item("MISSIGND").Enabled = True
                pForm.Items.Item("MISSCIES").Enabled = False
                pForm.Items.Item("MISESTSF").Enabled = False
                pForm.Items.Item("MISHARVP").Enabled = False
                pForm.Items.Item("MISAGETR").Enabled = False
                pForm.Items.Item("MISESTHD").Enabled = False
                pForm.Items.Item("MISESTLF").Enabled = False
                pForm.Items.Item("MISESTHQ").Enabled = False
                pForm.Items.Item("MISNFDIE").Enabled = False
                pForm.Items.Item("MISNETPUCD").Enabled = False
                pForm.Items.Item("MISGENCD").Enabled = False
                pForm.Items.Item("MISPROSR").Enabled = False
                pForm.Items.Item("MISHARVD").Enabled = False
                pForm.Items.Item("MISHARVQ").Enabled = False
                pForm.Items.Item("MISPROHR").Enabled = False

            End If

        ElseIf Project = "Harvest" Then
            Dim strSql As String
            Dim objRecSet As SAPbobsCOM.Recordset
            Dim initFishQty As Double
            Dim FeedConsumtionEstimate As Double


            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'strSql = "SELECT U_MISINIFQ = (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 " & _
            '"ON T1.DocEntry = T2.DocEntry AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '", U_MISFEEDQ = (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '"AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '",U_MISINIFC = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '"AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '",U_MISFEEDC = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '"AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '",U_MISNFDIE = (Select ISNULL(SUM(T1.U_MISFISHQ), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '"AND T2.U_MISTRXTP = 1 AND T1.U_MISINFO = '2' AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '",U_MISTPCST = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '"AND (T2.U_MISTRXTP = 1 OR T2.U_MISTRXTP = 2) AND T1.U_MISPROID = T0.U_MISPROID) " & _
            '"FROM [@MIS_PRJMSTR] T0 WHERE T0.U_MISNETST <> 'D' AND T0.U_MISPROID = '" & pForm.Items.Item("MISPROID").Specific.string & "'"


            'strSql = "SELECT U_MISINIFQ = (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '        "AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
            '        "(SELECT  ISNULL(SUM(G1.Quantity), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
            '        "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO')), " & _
            '        "U_MISFEEDQ = (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
            '        "(SELECT  ISNULL(SUM(G1.Quantity), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry  " & _
            '        "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE LEFT(ItemCode,2) = 'FE'), " & _
            '        "U_MISINIFC = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '        "AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
            '        "(SELECT  ISNULL(SUM(G1.Quantity * G1.StockPrice), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
            '        "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO')), " & _
            '        "U_MISFEEDC = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '        "AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
            '        "(SELECT  ISNULL(SUM(G1.Quantity * G1.StockPrice), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
            '        "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE LEFT(ItemCode,2) = 'FE'), " & _
            '        "U_MISNFDIE = (Select ISNULL(SUM(T1.U_MISFISHQ), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '        "AND T2.U_MISTRXTP = 1 AND T1.U_MISINFO = '2' AND T1.U_MISPROID = T0.U_MISPROID) , " & _
            '        "U_MISTPCST = (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
            '        "AND (T2.U_MISTRXTP = 1 OR T2.U_MISTRXTP = 2) AND T1.U_MISPROID = T0.U_MISPROID) - " & _
            '        "(Select ISNULL(SUM(G1.Quantity * G1.StockPrice), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
            '        "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO') OR LEFT(ItemCode,2) = 'FE') " & _
            '        "FROM [@MIS_PRJMSTR] T0 WHERE T0.U_MISNETST <> 'D' AND T0.U_MISPROID = '" & pForm.Items.Item("MISPROID").Specific.string & "' "


            'setelah menambah stock akhir

            strSql = "SELECT U_MISINIFQ = ISNULL(T0.U_MISINIFQ,0) + (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry  " & _
                    "AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
                    "(SELECT  ISNULL(SUM(G1.Quantity), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
                    "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO')), " & _
                    "U_MISFEEDQ = ISNULL(T0.U_MISFEEDQ,0) + (Select ISNULL(SUM(T1.Quantity), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
                    "(SELECT  ISNULL(SUM(G1.Quantity), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
                    "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE LEFT(ItemCode,2) = 'FE'), " & _
                    "U_MISINIFC = ISNULL(T0.U_MISINIFC,0) + (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
                    "AND T2.U_MISTRXTP = 2 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
                    "(SELECT  ISNULL(SUM(G1.Quantity * G1.Price), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
                    "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO')), " & _
                    "U_MISFEEDC = ISNULL(T0.U_MISFEEDC,0) + (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
                    "AND T2.U_MISTRXTP = 1 AND T1.U_MISPROID = T0.U_MISPROID) - " & _
                    "(SELECT  ISNULL(SUM(G1.Quantity * G1.Price), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
                    "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE LEFT(ItemCode,2) = 'FE'), " & _
                    "U_MISNFDIE = ISNULL(T0.U_MISNFDIE,0) + (Select ISNULL(SUM(T1.U_MISFISHQ), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
                    "AND T2.U_MISTRXTP = 1 AND T1.U_MISINFO = '2' AND T1.U_MISPROID = T0.U_MISPROID) , " & _
                    "U_MISTPCST = ISNULL(T0.U_MISTPCST,0) + (Select ISNULL(SUM(T1.Quantity * T1.StockPrice), 0) FROM IGE1 T1 INNER JOIN OIGE T2 ON T1.DocEntry = T2.DocEntry " & _
                    "AND (T2.U_MISTRXTP = 1 OR T2.U_MISTRXTP = 2) AND T1.U_MISPROID = T0.U_MISPROID) - " & _
                    "(Select ISNULL(SUM(G1.Quantity * G1.Price), 0) FROM IGN1 G1 INNER JOIN OIGN G2 ON G1.DocEntry = G2.DocEntry " & _
                    "AND (G2.U_MISTRXTP = 6) AND G1.U_MISPROID = T0.U_MISPROID WHERE (LEFT(ItemCode,2) = 'FL' OR LEFT(ItemCode,2) = 'GO') OR LEFT(ItemCode,2) = 'FE') " & _
                    "FROM [@MIS_PRJMSTR] T0 WHERE T0.U_MISNETST <> 'D' AND T0.U_MISPROID = '" & pForm.Items.Item("MISPROID").Specific.string & "' "



            objRecSet.DoQuery(strSql)

            If objRecSet.RecordCount > 0 Then
                initFishQty = objRecSet.Fields.Item("U_MISINIFQ").Value
                oFormProjectHarvest.Items.Item("MISINIFQ").Specific.string = initFishQty
                oFormProjectHarvest.Items.Item("MISFEEDQ").Specific.string = objRecSet.Fields.Item("U_MISFEEDQ").Value
                FeedConsumtionEstimate = (oFormProjectHarvest.Items.Item("MISFEEDQ").Specific.value * oFormProjectHarvest.Items.Item("MISFCR").Specific.Value) / 100
                oFormProjectHarvest.Items.Item("MISFCE").Specific.Value = FeedConsumtionEstimate
                oFormProjectHarvest.Items.Item("MISTEFQK").Specific.string = initFishQty + FeedConsumtionEstimate
                oFormProjectHarvest.Items.Item("MISINIFC").Specific.string = objRecSet.Fields.Item("U_MISINIFC").Value
                oFormProjectHarvest.Items.Item("MISFEEDC").Specific.string = objRecSet.Fields.Item("U_MISFEEDC").Value
                oFormProjectHarvest.Items.Item("MISNFDIE").Specific.string = objRecSet.Fields.Item("U_MISNFDIE").Value
                oFormProjectHarvest.Items.Item("MISTPCST").Specific.string = objRecSet.Fields.Item("U_MISTPCST").Value

                If oFormProjectHarvest.Items.Item("MISINIFQ").Specific.string <> 0 Or _
                oFormProjectHarvest.Items.Item("MISFEEDQ").Specific.string <> 0 Or _
                oFormProjectHarvest.Items.Item("MISINIFC").Specific.string <> 0 Or _
                oFormProjectHarvest.Items.Item("MISFEEDC").Specific.string <> 0 Or _
                oFormProjectHarvest.Items.Item("MISNFDIE").Specific.string <> 0 Or _
                oFormProjectHarvest.Items.Item("MISTPCST").Specific.string <> 0 Then
                    oFormProjectHarvest.Items.Item("MISPROCS").Specific.string = 1
                Else
                    oFormProjectHarvest.Items.Item("MISPROCS").Specific.string = 0
                End If
            Else
                objApplication.StatusBar.SetText("Transaction Closed ~17.0001~", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If




            If pForm.Items.Item("MISNETST").Specific.string <> "O" Then
                pForm.Items.Item("btnOK").Specific.Caption = "OK"
                pForm.Items.Item("MISNETID").Enabled = True
                pForm.Items.Item("MISNETID").Click()
                pForm.Items.Item("MISSIGND").Enabled = False
                pForm.Items.Item("MISSCIES").Enabled = False
                pForm.Items.Item("MISESTSF").Enabled = False
                pForm.Items.Item("MISHARVP").Enabled = False
                pForm.Items.Item("MISAGETR").Enabled = False
                pForm.Items.Item("MISESTHD").Enabled = False
                pForm.Items.Item("MISESTLF").Enabled = False
                pForm.Items.Item("MISESTHQ").Enabled = False
                pForm.Items.Item("MISNFDIE").Enabled = False
                pForm.Items.Item("MISNETPUCD").Enabled = False
                pForm.Items.Item("MISGENCD").Enabled = False
                pForm.Items.Item("MISPROSR").Enabled = False
                pForm.Items.Item("MISHARVD").Enabled = False
                pForm.Items.Item("MISHARVQ").Enabled = False
                pForm.Items.Item("MISPROHR").Enabled = False

            Else
                If pForm.Items.Item("btnOK").Specific.Caption = "FIND" Or pForm.Items.Item("btnOK").Specific.Caption = "OK" Then
                    pForm.Items.Item("btnOK").Specific.Caption = "HARVEST"
                End If
                pForm.Items.Item("MISHARVD").Enabled = True
                pForm.Items.Item("MISHARVQ").Enabled = True
                pForm.Items.Item("MISPROHR").Enabled = True
                pForm.Items.Item("MISHARVD").Click()
                pForm.Items.Item("MISNETID").Enabled = False
                pForm.Items.Item("MISSIGND").Enabled = False
                pForm.Items.Item("MISSCIES").Enabled = False
                pForm.Items.Item("MISESTSF").Enabled = False
                pForm.Items.Item("MISHARVP").Enabled = False
                pForm.Items.Item("MISAGETR").Enabled = False
                pForm.Items.Item("MISESTHD").Enabled = False
                pForm.Items.Item("MISESTLF").Enabled = False
                pForm.Items.Item("MISESTHQ").Enabled = False
                pForm.Items.Item("MISNFDIE").Enabled = False
                pForm.Items.Item("MISNETPUCD").Enabled = False
                pForm.Items.Item("MISGENCD").Enabled = False
                pForm.Items.Item("MISPROSR").Enabled = False
            End If
        End If

        pForm.Freeze(False)

        System.Runtime.InteropServices.Marshal.ReleaseComObject(pRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pForm)

    End Sub

End Class

