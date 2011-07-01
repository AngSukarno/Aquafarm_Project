Option Strict Off
Option Explicit On

Module SubMain
    Public Sub Main()
        Dim oClsConnection As clsConnection
        Dim oClsMain As clsMain
        'Dim oclsReproduksi As clsReproduksi
        'Dim oclsPromoRegister As clsPromoRegister
        'Dim oclsPromotion As clsPromotion
        'Dim oclsQCStatus As clsQCUpdate
        Dim strClass As String = String.Empty


        Try
            oClsConnection = New clsConnection
            strClass = "Main"
            oClsMain = New clsMain()


            'strClass = "Reproduksi"
            'oclsReproduksi = New clsReproduksi()
            'strClass = "Promo Register"
            'oclsPromoRegister = New clsPromoRegister()
            'strClass = "Promotion"
            'oclsPromotion = New clsPromotion
            'strClass = "QC Status"
            'oclsQCStatus = New clsQCUpdate


            'oClsMain.objApplication.StatusBar.SetText("AN AddOns Running.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            MsgBox(strClass & " failed.")
        End Try

    End Sub
End Module
