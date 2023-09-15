Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmReportViewer
    Private p_oReport As New ReportDocument

    Public Property ReportDocument() As ReportDocument
        Get
            Return p_oReport
        End Get
        Set(ByVal foValue As ReportDocument)
            p_oReport = foValue
        End Set
    End Property

    Private Sub frmReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        rptViewer.ReportSource = p_oReport
        rptViewer.Show()
    End Sub
End Class

