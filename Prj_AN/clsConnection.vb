Option Strict Off
Option Explicit On

Public Class clsConnection
    Public Shared objApplication As Object
    Public Shared objCompany As Object
    Public CCompany As SAPbobsCOM.Company
    Public WithEvents CApplication As SAPbouiCOM.Application


    Public Sub New()
        MyBase.New()

        Try
           
            SetApplication()

            '//*************************************************************
            '// Connect to DI
            '//*************************************************************

            CCompany = New SAPbobsCOM.Company

            '//get DI company (via UI)

            CCompany = CApplication.Company.GetDICompany()
            objCompany = CCompany



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub SetApplication()
        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String


        SboGuiApi = New SAPbouiCOM.SboGuiApi

        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode

        'sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

        sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs.GetValue(1))

        '// connect to a running SBO Application

        SboGuiApi.Connect(sConnectionString)

        '// get an initialized application object

        CApplication = SboGuiApi.GetApplication()
        objApplication = CApplication
    End Sub


End Class
