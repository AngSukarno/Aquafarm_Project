Option Strict Off
Option Explicit On
Public Class clsGlobal

    Const ListGRPO_FormId As String = "2000000001"
    Dim intFormCountListGRPO As Integer
    Dim strCurntListGRPO As String
    Dim blnModalListGRPO As Boolean
    Dim objFormListGRPO As SAPbouiCOM.Form



    Public Function fctFormatDate(ByVal pdate As Date, ByVal oCompany As SAPbobsCOM.Company, Optional ByVal sngFormat As Integer = 5) As String
        Dim strSeparator As String
        Dim oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        fctFormatDate = ""

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate
        strSeparator = oAdminInfo.DateSeparator

        Select Case sngFormat
            Case 0
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "yy")
            Case 1
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + "20" + Format(pdate, "yy")
            Case 2
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + Format(pdate, "yy")
            Case 3
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + "20" + Format(pdate, "yy")
            Case 4
                fctFormatDate = "20" + Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
            Case 5
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MMMM") + strSeparator + Format(pdate, "yyyy")
        End Select

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
    End Function

    Public Function fctSeparator(ByVal oCompany As SAPbobsCOM.Company, ByVal Value As String, Optional ByVal sngFormat As Integer = 5) As String
        Static oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing
        Dim DecimalSep As String
        Dim ThousandSep As String

        On Error GoTo ErrorHandler

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo


        If Value = "" Then
            GoTo ErrorHandler
        End If

        Select Case sngFormat
            Case 0
                fctSeparator = Left(Value, (Len(Value) - 3)) + "." + Right(Value, 2)
            Case 1
                fctSeparator = Left(Value, (Len(Value) - 3)) + "." + Right(Value, 2)
        End Select

        GoTo SetNothing

ErrorHandler:
        fctSeparator = ""

SetNothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)

    End Function

    Public Function fctFormatDateSave(ByVal oCompany As SAPbobsCOM.Company, ByVal pdate As String, ByVal sngFormat As Integer) As String
        Dim strFormat As String
        Dim strMonth As String
        Dim intLength As Integer
        Static oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        On Error GoTo ErrorHandler

        strMonth = "JANUARY01FEBRUARY02MARCH03APRIL04MAY05JUNE06JULY07AUGUST08SEPTEMBER09OCTOBER10NOVEMBER11DECEMBER12"

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate

        If pdate = "" Then
            GoTo ErrorHandler
        End If

        Select Case sngFormat
            Case 0
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 1
                fctFormatDateSave = Right(pdate, 4) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 2
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 3
                fctFormatDateSave = Right(pdate, 4) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 4
                fctFormatDateSave = Left(pdate, 4) + "/" + Mid(pdate, 6, 2) + "/" + Right(pdate, 2)
            Case 5
                intLength = InStr(1, strMonth, UCase(Mid(pdate, 4, Len(pdate) - 8))) + Len(Mid(pdate, 4, Len(pdate) - 8))
                fctFormatDateSave = Right(pdate, 4) + "/" + Mid(strMonth, intLength, 2) + "/" + Left(pdate, 2)
        End Select

        GoTo SetNothing

ErrorHandler:
        fctFormatDateSave = ""

SetNothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
        oGetCompanyService = Nothing
        oAdminInfo = Nothing
    End Function

    Public Sub subScrPaintGlobal(ByVal pFile As String, ByVal pFormId As String, _
                       ByRef pCounter As Integer, ByVal pForm As SAPbouiCOM.Form, ByVal RApplication As SAPbouiCOM.Application)

        Dim strScrPaintLoc As String
        Dim oXML As MSXML2.DOMDocument = Nothing

        strScrPaintLoc = Application.StartupPath & "\" & pFile

        oXML = New MSXML2.DOMDocument

        oXML.load(strScrPaintLoc)
        oXML.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
            oXML.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & pCounter

        pCounter = pCounter + 1

        RApplication.LoadBatchActions(oXML.xml)

        pForm = RApplication.Forms.GetForm(pFormId, 0)
        'oXML = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXML)
    End Sub

    Public Sub subAddCFLHandsetItem(ByVal oApplication As SAPbouiCOM.Application, ByVal strparam As String, ByVal struniqId As String)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCFL As SAPbouiCOM.ChooseFromList = Nothing
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams = Nothing
        Dim oCons As SAPbouiCOM.Conditions = Nothing
        Dim oCon As SAPbouiCOM.Condition = Nothing

        Try
            oForm = oApplication.Forms.ActiveForm
            oCFLs = oForm.ChooseFromLists

            oCFLCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"

            oCFLCreationParams.UniqueID = struniqId
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = strparam '"Itemcode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

SetNothing:
        'oCFLs = Nothing
        'oCFL = Nothing
        'oCFLCreationParams = Nothing
        'oForm = Nothing
        'oCons = Nothing
        'oCon = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon)
    End Sub

    Public Function fctFormExistGlobal(ByVal pFormId As String, ByRef pLoop As Integer, ByVal oApplication As SAPbouiCOM.Application) As Boolean
        Dim oForms As SAPbouiCOM.Forms = Nothing
        Dim intLoop As Integer

        fctFormExistGlobal = False
        pLoop = 0

        oForms = oApplication.Forms

        If oForms.Count > 0 Then
            For intLoop = 0 To oForms.Count - 1
                If oForms.Item(intLoop).Type = pFormId Then
                    fctFormExistGlobal = True
                    pLoop = intLoop
                    oForms = Nothing
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oForms)                    
                    Exit Function
                End If
            Next
        End If

        'oForms = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForms)
    End Function

    Public Function fctFindColourGlobal(ByVal pItemCode As String, ByVal oCompany As SAPbobsCOM.Company) As String
        Dim strSQL As String
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing
        Dim oRecSet1 As SAPbobsCOM.Recordset = Nothing

        Dim intCount As Integer

        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSet1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        fctFindColourGlobal = ""

        strSQL = "select "

        For intCount = 1 To 64
            strSQL = strSQL & "QryGroup" & intCount & ", "
        Next

        strSQL = Mid(strSQL, 1, Len(strSQL) - 2) & " from OITM where ItemCode = '" & pItemCode & "'"
        oRecSet.DoQuery(strSQL)
        If oRecSet.RecordCount <> 0 Then
            For intCount = 1 To 64
                If oRecSet.Fields.Item("QryGroup" & intCount).Value = "Y" Then
                    strSQL = "SELECT T0.[ItmsGrpNam] FROM OITG T0 WHERE T0.[ItmsTypCod] = '" & intCount & "'"
                    oRecSet1.DoQuery(strSQL)
                    fctFindColourGlobal = fctFindColourGlobal & oRecSet1.Fields.Item(0).Value & " "
                End If
            Next
        End If

        fctFindColourGlobal = Trim(fctFindColourGlobal)

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet1)
    End Function

    Public Function SetApplicationGlobal() As Object
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        SboGuiApi.Connect(sConnectionString)
        SetApplicationGlobal = SboGuiApi.GetApplication()



    End Function


    Public Sub AddMenuItemsGlobal(ByVal SAP_MenuId As String, ByVal AddOns_MenuId As String, ByVal AddOns_MenuDesc As String, ByVal Position_Menu As Integer, ByVal OAplication As Object) 'SAPbouiCOM.Application)
        Dim oMenus As SAPbouiCOM.Menus = Nothing
        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing

        'get the menu collection from the application
        If OAplication Is Nothing Then
            Dim SboGuiApi As SAPbouiCOM.SboGuiApi
            Dim sConnectionString As String

            SboGuiApi = New SAPbouiCOM.SboGuiApi

            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)
            OAplication = SboGuiApi.GetApplication()

        End If


        oMenus = OAplication.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams = Nothing
        oCreationPackage = OAplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

        'Kelompok : Production
        oMenuItem = OAplication.Menus.Item(SAP_MenuId)

        If Not oMenuItem.SubMenus.Exists(AddOns_MenuId) Then
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            With oCreationPackage
                .UniqueID = AddOns_MenuId
                .String = AddOns_MenuDesc
                .Enabled = True
                .Position = Position_Menu
            End With
            Try
                oMenuItem.SubMenus.AddEx(oCreationPackage)
            Catch ex As Exception
                Beep()
            End Try

        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenus)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage)
    End Sub

    Public Function fctCheckPostingDate(ByVal oCompany As SAPbobsCOM.Company, ByRef pBubbleEvent As Boolean, ByVal dtparam As String) As String
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing
        Dim strSQL As String
        Dim intResult As Integer

        pBubbleEvent = False
        fctCheckPostingDate = ""
        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strSQL = "SELECT top 1 T0.[Name], T0.[F_RefDate], T0.[T_RefDate], T0.[PeriodStat] FROM OFPR T0 WHERE T0.[PeriodStat] = 'N'"
        oRecSet.DoQuery(strSQL)
        If oRecSet.RecordCount <> 0 Then
            intResult = DateTime.Compare(CDate(fctFormatDateSave(oCompany, dtparam, 5)), oRecSet.Fields.Item("F_RefDate").Value)
            If intResult >= 0 Then
                intResult = DateTime.Compare(CDate(fctFormatDateSave(oCompany, dtparam, 5)), oRecSet.Fields.Item("T_RefDate").Value)
                If intResult <= 0 Then
                Else
                    fctCheckPostingDate = "Posting Date out of range."
                    GoTo SetNothing
                End If
            Else
                fctCheckPostingDate = "Posting Date out of range."
                GoTo SetNothing
            End If
        End If

        pBubbleEvent = True
SetNothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        oRecSet = Nothing
    End Function

End Class
