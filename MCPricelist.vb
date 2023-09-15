'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC Pricelist
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
'  iMac [ 2016.07.11 11:20 am ]
'       Translates clsMCPricelist by Sir Rex from VB6 to VB.net
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class MCPricelist
    Private Const p_sMsgHeadr As String = "MCPricelist"

    Private p_oApp As GRider
    Private p_oCashPrice As DataTable
    Private p_oInsPrice As DataTable

    Private p_sMCCatIDx As String
    Private p_sMCCatNme As String
    Private p_sModelIDx As String
    Private p_sModelNme As String

    Private p_nSelPrice As Double
    Private p_nLastPrce As Double
    Private p_nRebatesx As Double
    Private p_nMiscChrg As Double
    Private p_nMinDownx As Double
    Private p_nEndMrtgg As Double
    Private p_nAddPurc As Double

    Private p_bExactNm As Boolean
    Private p_bByCode As Boolean
    Private pbInitTran As Boolean

    Public Event CashPriceLoaded()

    Public ReadOnly Property CashPrice(ByVal Row As Integer, ByVal Index As Object) As Object
        Get
            If Not pbInitTran Then Return 0
            With p_oCashPrice
                If Row > (.Rows.Count - 1) Then Return 0
                Return .Rows(Row).Item(Index)
            End With
        End Get
    End Property

    Public ReadOnly Property InstallmentPrice(ByVal Row As Long, ByVal Index As Object) As Object
        Get
            If Not pbInitTran Then Return 0
            With p_oInsPrice
                If Row > (.Rows.Count - 1) Then Return 0

                If Index = "dInsPrice" Then
                    Return IIf(IFNull(.Rows(Row).Item("dInsPrice"), "2001-01-01") > IFNull(.Rows(Row).Item("xInsPrice"), "2001-01-01"), .Rows(Row).Item("dInsPrice"), IFNull(.Rows(Row).Item("xInsPrice"), "2001-01-01"))
                Else
                    Return .Rows(Row).Item(Index)
                End If
            End With
        End Get
    End Property

    Public Property AddtlLoan() As Double
        Get
            Return p_nAddPurc
        End Get
        Set(value As Double)
            p_nAddPurc = value
        End Set
    End Property

    Public Property MCCategory() As String
        Get
            Return p_sMCCatNme
        End Get
        Set(value As String)
            getCategory(value)
        End Set
    End Property

    Public Property MCModelID() As String
        Get
            Return p_sModelIDx
        End Get
        Set(value As String)
            p_bByCode = True
            getModel(value)
        End Set
    End Property

    Public Property MCModel() As String
        Get
            Return p_sModelNme
        End Get
        Set(value As String)
            p_bByCode = False
            getModel(value)
        End Set
    End Property

    Public ReadOnly Property Rebate() As Double
        Get
            Return p_nRebatesx
        End Get
    End Property

    Public ReadOnly Property MinimumDown() As Double
        Get
            Return p_nMinDownx
        End Get
    End Property

    Public ReadOnly Property MiscCharge() As Double
        Get
            Return p_nMiscChrg
        End Get
    End Property

    Public ReadOnly Property EndMortgage() As Double
        Get
            Return p_nEndMrtgg
        End Get
    End Property

    Public ReadOnly Property LastPrice() As Double
        Get
            Return p_nLastPrce
        End Get
    End Property

    Public ReadOnly Property SelPrice() As Double
        Get
            Return p_nSelPrice
        End Get
    End Property

    Public ReadOnly Property CashPriceCount() As Integer
        Get
            If Not pbInitTran Then Return 0

            Return p_oCashPrice.Rows.Count
        End Get
    End Property

    Public ReadOnly Property InstallmentPriceCount() As Integer
        Get
            If Not pbInitTran Then Return 0

            Return p_oInsPrice.Rows.Count
        End Get
    End Property

    Public ReadOnly Property CashLatestDate() As Date
        Get
            Dim ldLatest As Date
            Dim lnCtr As Integer

            With p_oCashPrice
                If .Rows.Count > 1 Then
                    ldLatest = IFNull(.Rows(0).Item("dPricexxx"), "2001-01-01")

                    For lnCtr = 0 To .Rows.Count - 1
                        If ldLatest < IFNull(.Rows(lnCtr).Item("dPricexxx"), "2001-01-01") Then ldLatest = IFNull(.Rows(0).Item("dPricexxx"), "2001-01-01")
                    Next
                    Return ldLatest
                Else
                    Return CDate("2001-01-01")
                End If
            End With
        End Get
    End Property

    Public ReadOnly Property InstallmentLatestDate() As Date
        Get
            Dim ldLatest As Date
            Dim lnCtr As Integer

            With p_oInsPrice
                If .Rows.Count > 1 Then
                    ldLatest = IIf(IFNull(.Rows(0).Item("dInsPrice"), "2001-01-01") _
                                    > IFNull(.Rows(0).Item("xInsPrice"), "2001-01-01"), _
                                        IFNull(.Rows(0).Item("dInsPrice"), "2001-01-01"), _
                                        IFNull(.Rows(0).Item("xInsPrice"), "2001-01-01"))

                    For lnCtr = 0 To .Rows.Count - 1
                        If ldLatest < IFNull(.Rows(lnCtr).Item("dInsPrice"), "2001-01-01") Then ldLatest = IFNull(.Rows(lnCtr).Item("dInsPrice"), "2001-01-01")
                        If ldLatest < IFNull(.Rows(lnCtr).Item("xInsPrice"), "2001-01-01") Then ldLatest = IFNull(.Rows(lnCtr).Item("xInsPrice"), "2001-01-01")
                    Next
                    Return ldLatest
                Else
                    Return CDate("2001-01-01")
                End If
            End With
        End Get
    End Property

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider

        pbInitTran = True
    End Sub

    Sub LoadCashPrice()
        Dim lsSQL As String

        If Not pbInitTran Then GoTo endProc

        If p_sMCCatIDx <> "" Then
            lsSQL = AddCondition(getSQLCash, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
        ElseIf p_sModelIDx <> "" Then
            lsSQL = AddCondition(getSQLCash, "a.sModelIDx = " & strParm(p_sModelIDx))
        Else
            lsSQL = getSQLCash()
        End If

        p_oCashPrice = p_oApp.ExecuteQuery(lsSQL)
        If p_oCashPrice.Rows.Count = 0 Then p_oCashPrice.Rows.Add()

        p_bExactNm = False
endProc:
        RaiseEvent CashPriceLoaded()
        Exit Sub
    End Sub

    Sub LoadInstallmentPrice()
        Dim lsSQL As String

        If pbInitTran = False Then GoTo endProc

        lsSQL = getSQLInstallment()
        If p_sMCCatIDx <> "" Then
            lsSQL = AddCondition(getSQLInstallment, "a.sMCCatIDx = " & strParm(p_sMCCatIDx))
        End If

        p_oInsPrice = p_oApp.ExecuteQuery(lsSQL)
        If p_oInsPrice.Rows.Count = 0 Then p_oInsPrice.Rows.Add()

        p_bExactNm = True
endProc:
        Exit Sub
    End Sub

    Function getMonthly(ByVal DownPayment As Double, Term As Integer) As Double
        Dim loDT As DataTable
        Dim lsSQL As String

        If p_sModelIDx = "" Then
            MsgBox("No Model has been Selected!" & vbCrLf & _
                     "Please select a model first then try again!", vbCritical, "Warning")
            Return 0
        End If

        lsSQL = "SELECT" & _
               "  a.nSelPrice" & _
               ", a.nMinDownx" & _
               ", b.nMiscChrg" & _
               ", b.nRebatesx" & _
               ", b.nEndMrtgg" & _
               ", c.nAcctThru" & _
               ", c.nFactorRt" & _
            " FROM MC_Model_Price a" & _
               ", MC_Category b" & _
               ", MC_Term_Category c" & _
            " WHERE a.sMCCatIDx = b.sMCCatIDx" & _
               " AND a.sMCCatIDx = c.sMCCatIDx" & _
               " AND a.sModelIDx = " & strParm(p_sModelIDx) & _
               " AND c.nAcctThru = " & Term

        loDT = p_oApp.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Return 0

        With loDT
            If Term < 4 Then
                Return Math.Round((.Rows(0).Item("nSelPrice") + p_nAddPurc - (DownPayment)) _
                                  * .Rows(0).Item("nFactorRt") / Term, 0)
            Else
                Return Math.Round(((.Rows(0).Item("nSelPrice") + p_nAddPurc - DownPayment + .Rows(0).Item("nMiscChrg")) _
                                   * .Rows(0).Item("nFactorRt") / Term) + .Rows(0).Item("nRebatesx") + (.Rows(0).Item("nEndMrtgg") / Term), 0)
            End If
        End With
    End Function

    Private Sub getCategory(ByVal lsValue As String)
        Dim lsSQL As String
        Dim loDT As DataTable

        If lsValue = p_sMCCatNme And lsValue <> "" Then GoTo endProc

        lsSQL = "SELECT" & _
                    "  sMCCatIDx" & _
                    ", sMCCatNme" & _
                 " FROM MC_Category" & _
                 " WHERE sMCCatNme LIKE " & strParm(Trim(lsValue) & "%")

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 1 Then
            p_sMCCatIDx = loDT(0)("sMCCatIDx")
            p_sMCCatNme = loDT(0)("sMCCatNme")
        Else
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                            , lsSQL _
                                            , True _
                                            , lsValue _
                                            , "sMCCatIDx»sMCCatNme" _
                                            , "ID»Category", _
                                            , "sMCCatIDx»sMCCatNme" _
                                            , 1)

            If IsNothing(loRow) Then
                GoTo endwithClear
            Else
                p_sMCCatIDx = loRow.Item("sMCCatIDx")
                p_sMCCatNme = loRow.Item("sMCCatNme")
            End If
        End If
endProc:
        loDT = Nothing

        Exit Sub
endwithClear:
        p_sMCCatIDx = ""
        p_sMCCatNme = ""
    End Sub

    Private Sub getModel(ByVal lsValue As String)
        Dim lsSQL As String
        Dim loDT As DataTable

        If p_bByCode Then
            If lsValue = p_sModelIDx Then GoTo endProc
            lsSQL = "a.sModelIDx = " & strParm(Trim(lsValue))
        Else
            If lsValue = p_sModelNme Then GoTo endProc
            If p_bExactNm Then
                lsSQL = "a.sModelNme = " & strParm(Trim(lsValue))
            Else
                lsSQL = "a.sModelNme LIKE " & strParm(Trim(lsValue) & "%")
            End If
        End If

        lsSQL = "SELECT" & _
               "  a.sModelIDx" & _
               ", a.sModelNme" & _
               ", b.nMinDownx" & _
               ", c.nRebatesx" & _
               ", c.nMiscChrg" & _
               ", c.nEndMrtgg" & _
               ", b.nSelPrice" & _
               ", b.nLastPrce" & _
            " FROM MC_Model a" & _
               ", MC_Model_Price b" & _
               ", MC_Category c" & _
            " WHERE a.sModelIDx = b.sModelIDx" & _
               " AND b.cRecdStat = " & strParm(xeRecordStat.RECORD_NEW) & _
               " AND b.sMCCatIDx = c.sMCCatIDx" & _
               IIf(lsSQL <> "", " AND " & lsSQL, "")

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 1 Then
            p_sModelIDx = loDT(0)("sModelIDx")
            p_sModelNme = loDT(0)("sModelNme")
            p_nRebatesx = loDT(0)("nRebatesx")
            p_nMiscChrg = loDT(0)("nMiscChrg")
            p_nEndMrtgg = loDT(0)("nEndMrtgg")
            p_nMinDownx = IIf(IsDBNull(loDT(0)("nMinDownx")), 0, loDT(0)("nMinDownx"))
            p_nSelPrice = loDT(0)("nSelPrice")
            p_nLastPrce = loDT(0)("nLastPrce")
        Else
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                            , lsSQL _
                                            , True _
                                            , lsValue _
                                            , "sModelIDx»sModelNme" _
                                            , "ID»Model", _
                                            , "sModelIDx»sModelNme" _
                                            , 1)

            If IsNothing(loRow) Then
                p_sModelIDx = ""
                p_sModelNme = ""
            Else
                p_sModelIDx = loRow.Item("sModelIDx")
                p_sModelNme = loRow.Item("sModelNme")
                p_nRebatesx = loRow.Item("nRebatesx")
                p_nMiscChrg = loRow.Item("nMiscChrg")
                p_nEndMrtgg = loRow.Item("nEndMrtgg")
                p_nMinDownx = IIf(IsDBNull(loRow.Item("nMinDownx")), 0, loRow.Item("nMinDownx"))
                p_nSelPrice = loRow.Item("nSelPrice")
                p_nLastPrce = loRow.Item("nLastPrce")
            End If
        End If
endProc:
        p_nAddPurc = 0
        loDT = Nothing

        Exit Sub
endwithClear:
        p_sModelIDx = ""
        p_sModelNme = ""
        p_nRebatesx = 0
        p_nMiscChrg = 0
        p_nEndMrtgg = 0

        GoTo endProc
    End Sub

    Private Function getSQLCash() As String
        Return "SELECT" & _
                     "  c.sMCCatNme" & _
                     ", b.sModelNme" & _
                     ", d.sBrandNme" & _
                     ", a.nSelPrice" & _
                     ", a.nLastPrce" & _
                     ", a.nDealrPrc" & _
                     ", a.dPricexxx" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                     ", MC_Category c" & _
                     ", Brand d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sMCCatIDx = c.sMCCatIDx" & _
                     " AND b.sBrandIDx = d.sBrandIDx" & _
                     " AND a.cRecdStat = " & strParm(xeRecordStat.RECORD_NEW) & _
                     " AND b.cRecdStat = " & strParm(xeRecordStat.RECORD_NEW) & _
                  " ORDER BY c.sMCCatNme, b.sModelNme"
    End Function

    Private Function getSQLInstallment() As String
        Return "SELECT" & _
                     "  c.sMCCatNme" & _
                     ", b.sModelNme" & _
                     ", a.nSelPrice" & _
                     ", a.nMinDownx" & _
                     ", d.nAcctThru" & _
                     ", d.nFactorRt" & _
                     ", a.dInsPrice" & _
                     ", d.dPricexxx xInsPrice" & _
                  " FROM MC_Model_Price a" & _
                     ", MC_Model b" & _
                     ", MC_Category c" & _
                     ", MC_Term_Category d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sMCCatIDx = c.sMCCatIDx" & _
                     " AND a.sMCCatIDx = d.sMCCatIDx" & _
                     " AND a.cRecdStat = " & strParm(xeRecordStat.RECORD_NEW) & _
                     " AND b.cRecdStat = " & strParm(xeRecordStat.RECORD_NEW) & _
                  " ORDER BY c.sMCCatNme, b.sModelNme"
    End Function
End Class
