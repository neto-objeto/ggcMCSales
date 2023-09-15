'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Report Viewer Object
'
' Copyright 2012 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Kalyptus [ 07/10/2013 09:13 am ]
'      Started creating of this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports CrystalDecisions.CrystalReports.Engine
Imports ggcAppDriver


Public Class Reports
    Private p_sRptIDx As String
    Private p_sHeadr1 As String
    Private p_sHeadr2 As String

    Private p_sProdctID As String
    Private p_nUserRght As Int32
    Private p_cRegularx As String
    Private p_cLogRepxx As String

    Private p_oFormxx As frmReportViewer
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oReport As ReportDocument

    Public Property GRider() As GRider
        Get
            Return p_oDriver
        End Get
        Set(ByVal foValue As GRider)
            p_oDriver = foValue
        End Set
    End Property

    Public Property Header1() As String
        Get
            Return p_sHeadr1
        End Get
        Set(ByVal fsValue As String)
            p_sHeadr1 = fsValue
        End Set
    End Property

    Public Property Header2() As String
        Get
            Return p_sHeadr2
        End Get
        Set(ByVal fsValue As String)
            p_sHeadr2 = fsValue
        End Set
    End Property

    Public Property ReportSource() As ReportDocument
        Get
            Return p_oReport
        End Get
        Set(ByVal foValue As ReportDocument)
            p_oReport = foValue
        End Set
    End Property

    Public Function initReport(ByVal fsReportID As String) As Boolean

        p_sRptIDx = fsReportID

        p_oReport = New ReportDocument
        'p_oReport.Load(p_oDriver.AppPath & "\vb.net\Reports\" & p_sRptIDx & ".rpt")
        'Return True

        Dim lsSQL As String
        Dim loDta As DataTable

        lsSQL = "SELECT" & _
                    "  sReportNm" & _
                    ", sFileName" & _
                    ", sProdctID" & _
                    ", nUserRght" & _
                    ", cRegularx" & _
                    ", cLogRepxx" & _
                 " FROM xxxReportOther" & _
                 " WHERE sReportID = " & strParm(p_sRptIDx)
        loDta = p_oDriver.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            MsgBox("Unable to find the Report Information for this Report ID!")
            Return False
        End If

        p_sProdctID = loDta(0).Item("sProdctID")
        p_nUserRght = loDta(0).Item("nUserRght")
        p_cRegularx = loDta(0).Item("cRegularx")
        p_cLogRepxx = loDta(0).Item("cLogRepxx")

        p_oReport.Load(p_oDriver.AppPath & "\vb.net\Reports\" & loDta(0).Item("sFileName") & ".rpt")
        Return True

    End Function

    Public Sub showReport()

        Dim loTxtObj As CrystalDecisions.CrystalReports.Engine.TextObject

        If p_cRegularx = "1" Then
            'Set Branch Name
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtCompany")
            loTxtObj.Text = p_oDriver.BranchName

            'Set Branch Address
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtAddress")
            loTxtObj.Text = p_oDriver.Address

            'Set First Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading1")
            loTxtObj.Text = p_sHeadr1

            'Set Second Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
            loTxtObj.Text = p_sHeadr2

            'Set Second Header
            loTxtObj = p_oReport.ReportDefinition.Sections(4).ReportObjects("txtRptUser")
            loTxtObj.Text = Decrypt(p_oDriver.UserName, "08220326")
        End If

        'p_oReport.PrintToPrinter(1, False, 0, 0)

        'Assigned Report Document to the CRViewer attached to p_oFormxx
        p_oFormxx.ReportDocument = p_oReport

        'Show the form where CRViewer was attached
        showModalForm(p_oFormxx, p_oDriver.MDI)
    End Sub

    Public Sub New()
        p_oReport = Nothing
        p_oFormxx = New frmReportViewer
    End Sub
End Class
