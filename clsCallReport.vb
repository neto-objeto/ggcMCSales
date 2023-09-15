'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     CR Request Monitor/Status Object
'
' Copyright 2015 and Beyond
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
'  Kalyptus [ 08/11/2015 11:00 am ]
'      Started creating of this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsCallReport
    Private p_oDriver As ggcAppDriver.GRider
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    Private p_nReptType As Integer      '0=Summary;1=Detail
    Private p_abInclude(1) As Integer   '0=Default;1=Financing
    Private p_dDateFrom As Date
    Private p_dDateThru As Date
    Private p_sBranchCD As String

    Public Function getParameter() As Boolean
        Dim loFrm As frmOutboundCall

        loFrm = New frmOutboundCall

        'Disable Report Type Group
        loFrm.gbxPanel01.Enabled = False
        'Set Detail as Report Type
        loFrm.rbtTypex02.Checked = True

        loFrm.ShowDialog()

        If loFrm.isOkey Then
            'Since we have not allowed the report type to be edited
            p_nReptType = 0

            p_abInclude(0) = loFrm.chkInclude01.Checked
            p_abInclude(1) = loFrm.chkInclude02.Checked

            p_dDateFrom = loFrm.txtField01.Text
            p_dDateThru = loFrm.txtField02.Text

            loFrm = Nothing
            Return True
        Else
            loFrm = Nothing
            Return False
        End If
    End Function

    Public Function ReportTrans() As Boolean

        'she 2016-05-16 1:37pm
        Dim lsRqstType As String
        Dim oProg As frmProgress

        'she 2016-05-14
        Dim lsSQL As String 'whole statement
        Dim lsQuery1 As String
        Dim lsQuery2 As String

        'Show progress bar
        oProg = New frmProgress
        oProg.PistonInfo = p_oDriver.AppPath & "/piston.avi"
        oProg.ShowTitle("EXTRACTING RECORDS FROM DATABASE")
        oProg.ShowProcess("Please wait...")
        oProg.Show()

        lsQuery1 = "SELECT b.sCompnyNm `sClientNm`" & _
                        ", a.sMobileNo" & _
                        ", a.sRemarksx" & _
                        ", a.sApprovCd" & _
                        ", a.cTLMStatx" & _
                " FROM Call_Outgoing a" & _
                    " LEFT JOIN Client_Master b" & _
                        " ON a.sClientID = b.sClientID" & _
                " WHERE a.sSourceCd <> 'LEND'" & _
                    " AND a.cTranStat = '2'" & _
                    " AND a.sModified = " & strParm(p_oDriver.UserID) & _
                    " AND a.dModifled LIKE " & strParm(Format(p_oDriver.SysDate, "yyyy-MM-dd")) & _
                " ORDER BY a.dModified ASC"

        lsQuery2 = "SELECT b.sCompnyNm `sClientNm`" & _
                       ", a.sMobileNo" & _
                       ", a.sRemarksx" & _
                       ", a.sApprovCd" & _
                       ", a.cTLMStatx" & _
               " FROM Call_Outgoing a" & _
                   " LEFT JOIN Client_Master b" & _
                       " ON a.sClientID = b.sClientID" & _
               " WHERE a.sSourceCd = 'LEND'" & _
                   " AND a.cTranStat = '2'" & _
                   " AND a.sModified = " & strParm(p_oDriver.UserID & _
                   " AND a.dModifled LIKE " & strParm(Format(p_oDriver.SysDate, "yyyy-MM-dd"))) & _
                " ORDER BY a.dModified ASC"

        lsSQL = ""
        If p_abInclude(0) Then
            lsSQL = lsQuery1
        End If

        If p_abInclude(1) And lsSQL <> "" Then
            lsSQL = lsSQL & " UNION " & lsQuery2
        ElseIf p_abInclude(1) Then
            lsSQL = lsQuery2
        End If

        Debug.Print(lsSQL)

        p_oDTSrce = p_oDriver.ExecuteQuery(lsSQL)

        Dim loDtaTbl As DataTable = getRptTable()
        Dim lnCtr As Integer

        oProg.ShowTitle("LOADING RECORDS")
        oProg.MaxValue = p_oDTSrce.Rows.Count

        For lnCtr = 0 To p_oDTSrce.Rows.Count - 1

            oProg.ShowProcess("Loading " & p_oDTSrce(lnCtr).Item("sClientNm") & "...")

            loDtaTbl.Rows.Add(addRow(lnCtr, loDtaTbl))
        Next

        oProg.ShowSuccess()

        Dim clsRpt As ggcTLMReport.Reports
        clsRpt = New ggcTLMReport.Reports
        clsRpt.GRider = p_oDriver
        'Set the Report Source Here
        If Not clsRpt.initReport("TLMC1") Then
            Return False
        End If

        Dim loRpt As ReportDocument = clsRpt.ReportSource

        'Dim sec As CrystalDecisions.CrystalReports.Engine.Section
        'For Each sec In loRpt.ReportDefinition.Sections
        'MsgBox(sec.Name)
        'Next

        Dim loTxtObj As CrystalDecisions.CrystalReports.Engine.TextObject
        loTxtObj = loRpt.ReportDefinition.Sections(0).ReportObjects("txtCompany")
        loTxtObj.Text = p_oDriver.BranchName

        'Set Branch Address
        loTxtObj = loRpt.ReportDefinition.Sections(0).ReportObjects("txtAddress")
        loTxtObj.Text = p_oDriver.Address

        'Set First Header
        loTxtObj = loRpt.ReportDefinition.Sections(1).ReportObjects("txtHeading1")
        loTxtObj.Text = "Outbound Call Report"

        'Set Second Header
        loTxtObj = loRpt.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
        loTxtObj.Text = "(" & Format(p_dDateFrom, "yyyy-MM-dd") & " TO " & Format(p_dDateThru, "yyyy-MM-dd") & ")"

        loTxtObj = loRpt.ReportDefinition.Sections(10).ReportObjects("txtRptUser")
        loTxtObj.Text = Decrypt(p_oDriver.UserName, "08220326")

        loRpt.SetDataSource(p_oSTRept)
        clsRpt.showReport()

        Return True
    End Function

    Private Function getRptTable() As DataTable
        'Initialize DataSet
        p_oSTRept = New DataSet

        'Load the data structure of the Dataset
        'Data structure was saved at DataSet1.xsd 
        p_oSTRept.ReadXmlSchema(p_oDriver.AppPath & "\vb.net\Reports\DataSet1.xsd")

        'Return the schema of the datatable derive from the DataSet 
        Return p_oSTRept.Tables(0)
    End Function

    Private Function addRow(ByVal lnRow As Integer, ByVal foSchemaTable As DataTable) As DataRow
        'ByVal foDTInclue As DataTable
        Dim loDtaRow As DataRow

        'Create row based on the schema of foSchemaTable
        loDtaRow = foSchemaTable.NewRow

        loDtaRow.Item("sField01") = p_oDTSrce(lnRow).Item("sClientNm")
        loDtaRow.Item("sField02") = p_oDTSrce(lnRow).Item("sMobileNo")
        loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("sApprovCd")
        loDtaRow.Item("sField04") = p_oDTSrce(lnRow).Item("cTLMStatx")
        loDtaRow.Item("sField05") = p_oDTSrce(lnRow).Item("sRemarksx")

        Return loDtaRow
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oDriver = foRider
        p_oSTRept = Nothing
        p_oDTSrce = Nothing
    End Sub
End Class