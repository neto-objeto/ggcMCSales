'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Telemarketing Classification
'
' Copyright 2014 and Beyond
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
'  Jheff [ 05/29/2014 04:27 pm ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports ggcAppDriver

Public Class TLMCustClassify
    Private Const pxeMODULENAME As String = "clsTLMCustClassify"
    Private p_oMaster As DataTable

    Private p_oAppDrvr As GRider

    WriteOnly Property AppDriver() As GRider
        Set(ByVal oAppDriver As GRider)
            p_oAppDrvr = oAppDriver
        End Set
    End Property

    Function ClassifyMPCustomer() As Boolean
        Dim lsProcName As String
        Dim lsClientID As String = ""
        Dim lbRecExist As Boolean
        Dim lnRow As Integer = 0

        lsProcName = "ClassifyMPCustomer"
        'On Error GoTo errProc

        If Not getMPRecords() Then
            Return False
            GoTo endProc
        End If

        'p_oAppDrvr.BeginTransaction()
        With p_oMaster
            For Each drRow As DataRow In .Rows
                ' set to default value
                lbRecExist = Not IsDBNull(drRow("xClientID"))

                If lbRecExist Then
                    If Not updateMPTLMClass(lnRow) Then
                        ClassifyMPCustomer = False
                        Return False
                        GoTo endProc
                    End If
                Else
                    If lsClientID <> drRow("sClientID") Then
                        If addMPTLM(lnRow) = False Then
                            Return False
                            GoTo endProc
                        End If
                    Else
                        If updateCPSales(lnRow) = False Then
                            Return False
                            GoTo endProc
                        End If
                    End If
                End If

                lsClientID = drRow("sClientID")
                lnRow = lnRow + 1
            Next drRow
        End With
        'p_oAppDrvr.CommitTransaction()

        Return True

endProc:
        Exit Function
errProc:
        p_oAppDrvr.RollBackTransaction()
        ShowError(lsProcName & "( " & " ) ")
    End Function

    Function ClassifyMCCustomer(ByVal bCash As Boolean) As Boolean
        Dim lsProcName As String
        Dim lsClientID As String = ""
        Dim lbRecExist As Boolean
        Dim lnRow As Integer = 0

        lsProcName = "ClassifyMCCustomer"
        'On Error GoTo errProc

        If Not bCash Then
            If Not getMCRecords() Then
                Return False
                GoTo endProc
            End If
        Else
            If Not getMCSalesRecords() Then
                Return False
                GoTo endProc
            End If
        End If

        p_oAppDrvr.BeginTransaction()
        With p_oMaster
            For Each drRow As DataRow In .Rows
                ' set to default value
                lbRecExist = Not IsDBNull(drRow("xClientID"))

                If lbRecExist Then
                    If Not updateMCTLMClass(bCash, lnRow) Then
                        ClassifyMCCustomer = False
                        Return False
                        GoTo endProc
                    End If
                Else
                    If lsClientID <> drRow("sClientID") Then
                        If addMCTLM(bCash, lnRow) = False Then
                            Return False
                            GoTo endProc
                        End If
                    Else
                        If Not bCash Then
                            If updateARMaster(lnRow) = False Then
                                Return False
                                GoTo endProc
                            End If
                        Else
                            If updateSOMaster(lnRow) = False Then
                                Return False
                                GoTo endProc
                            End If
                        End If
                    End If
                End If

                lsClientID = drRow("sClientID")
                lnRow = lnRow + 1
            Next drRow
        End With
        p_oAppDrvr.CommitTransaction()

        Return True

endProc:
        Exit Function
errProc:
        p_oAppDrvr.RollBackTransaction()
        ShowError(lsProcName & "( " & " ) ")
    End Function

    Private Function getMPRecords() As Boolean
        Dim lsProcName As String
        Dim lsSQL As String

        lsProcName = "getMPRecords"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", CONCAT(b.sLastName, ', ', b.sFrstName) xCustName" & _
                    ", CONCAT(b.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) xAddressx" & _
                    ", b.dBirthDte" & _
                    ", (YEAR(CURDATE())-YEAR(b.dBirthDte))- (RIGHT(CURDATE(),5)<RIGHT(b.dBirthDte,5)) xAgexxxxx" & _
                    ", b.cGenderCd" & _
                    ", f.sOccptnNm" & _
                    ", b.sEmailAdd" & _
                    ", c.sClientID xClientID" & _
                    ", c.sClassIDx" & _
                    ", a.nCashAmtx" & _
                    ", g.nTranTotl xCCardAmt" & _
                    ", a.sTransNox" & _
                    ", a.cTranStat" & _
                    ", e.sProvIDxx" & _
                    ", c.cSourceCd" & _
                    ", c.dLastCall"
        lsSQL = lsSQL & _
                    ", a.sTransNox" & _
                 " FROM CP_SO_Master a" & _
                    " LEFT JOIN CP_SO_Credit_Card g" & _
                       " ON a.sTransNox = g.sTransNox" & _
                    " LEFT JOIN TLM_Client c" & _
                       " ON a.sClientID = c.sClientID" & _
                    ", Client_Master b" & _
                       " LEFT JOIN TownCity d" & _
                          " LEFT JOIN Province e" & _
                             " ON d.sProvIDxx = e.sProvIDxx" & _
                          " ON b.sTownIDxx = d.sTownIDxx" & _
                       " LEFT JOIN Occupation f" & _
                          " ON b.sOccptnID = f.sOccptnID" & _
                 " WHERE a.sClientID = b.sClientID" & _
                    " AND a.dTransact > '2012-01-01'" & _
                    " AND NOT (a.cTranStat = " & strParm(xeTranStat.TRANS_CANCELLED) & _
                       " OR a.cTranStat >= '4')" & _
                    " AND (LENGTH(b.sMobileNo) = 11" & _
                       " OR LENGTH(b.sPhoneNox) = 11)" & _
                 " ORDER BY a.sClientID, (a.nCashAmtx + IFNULL(g.nTranTotl,0)) DESC LIMIT 500000"
        p_oMaster = New DataTable
        Try
            p_oMaster = New DataTable
            p_oMaster = p_oAppDrvr.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
        End Try

        Return p_oMaster.Rows.Count > 0

endProc:
        Exit Function
errProc:
        ShowError(lsProcName & "( " & " ) ")
    End Function

    Private Function getMCRecords() As Boolean
        Dim lsProcName As String
        Dim lsSQL As String

        lsProcName = "getMCRecords"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", CONCAT(b.sLastName, ', ', b.sFrstName) xCustName" & _
                    ", CONCAT(b.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) xAddressx" & _
                    ", b.dBirthDte" & _
                    ", (YEAR(CURDATE())-YEAR(b.dBirthDte))- (RIGHT(CURDATE(),5)<RIGHT(b.dBirthDte,5)) xAgexxxxx" & _
                    ", b.cGenderCd" & _
                    ", f.sOccptnNm" & _
                    ", b.sEmailAdd" & _
                    ", c.sClientID xClientID" & _
                    ", c.sClassIDx" & _
                    ", a.sAcctNmbr" & _
                    ", e.sProvIDxx" & _
                    ", c.cSourceCd" & _
                    ", c.dLastCall"
        lsSQL = lsSQL & _
                 " FROM MC_AR_Master a" & _
                    " LEFT JOIN TLM_Client c" & _
                       " ON a.sClientID = c.sClientID" & _
                    ", Client_Master b" & _
                       " LEFT JOIN TownCity d" & _
                          " LEFT JOIN Province e" & _
                             " ON d.sProvIDxx = e.sProvIDxx" & _
                          " ON b.sTownIDxx = d.sTownIDxx" & _
                       " LEFT JOIN Occupation f" & _
                          " ON b.sOccptnID = f.sOccptnID" & _
                    ", MC_Credit_Application g" & _
                 " WHERE a.sClientID = b.sClientID" & _
                    " AND a.sApplicNo = g.sTransNox" & _
                    " AND ((g.cWithFinx = " & strParm(xeLogical.YES) & _
                       " AND a.dPurchase > '2012-12-31')" & _
                       " OR (a.cAcctStat = '1'" & _
                       " AND a.cRatingxx IN ('x', 'g'))" & _
                       " AND a.dClosedxx > '2012-12-31')" & _
                    " AND a.cPostedxx = " & strParm(xeLogical.NO) & _
                    " AND (LENGTH(b.sMobileNo) = 11" & _
                       " OR LENGTH(b.sPhoneNox) = 11)" & _
                    " AND (YEAR(CURDATE())-YEAR(b.dBirthDte)) - (RIGHT(CURDATE(),5)<RIGHT(b.dBirthDte,5)) >= 20" & _
                 " ORDER BY a.sClientID LIMIT 100000"

        p_oMaster = New DataTable
        Try
            p_oMaster = New DataTable
            p_oMaster = p_oAppDrvr.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
        End Try

        Return p_oMaster.Rows.Count > 0

endProc:
        Exit Function
errProc:
        ShowError(lsProcName & "( " & " ) ")
    End Function

    Private Function getMCSalesRecords() As Boolean
        Dim lsProcName As String
        Dim lsSQL As String

        lsProcName = "getMCRecords"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", CONCAT(b.sLastName, ', ', b.sFrstName) xCustName" & _
                    ", CONCAT(b.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) xAddressx" & _
                    ", b.dBirthDte" & _
                    ", (YEAR(CURDATE())-YEAR(b.dBirthDte))- (RIGHT(CURDATE(),5)<RIGHT(b.dBirthDte,5)) xAgexxxxx" & _
                    ", b.cGenderCd" & _
                    ", f.sOccptnNm" & _
                    ", b.sEmailAdd" & _
                    ", c.sClientID xClientID" & _
                    ", c.sClassIDx" & _
                    ", a.sTransNox" & _
                    ", e.sProvIDxx" & _
                    ", c.cSourceCd" & _
                    ", a.sTransNox" & _
                    ", g.sAcctNmbr" & _
                    ", c.dLastCall"

        lsSQL = lsSQL & _
                 " FROM MC_SO_Master a" & _
                    " LEFT JOIN MC_AR_Master g" & _
                       " ON a.dTransact = g.dPurchase" & _
                          " AND a.sClientID = g.sClientID" & _
                    " LEFT JOIN TLM_Client c" & _
                       " ON a.sClientID = c.sClientID" & _
                    ", Client_Master b" & _
                       " LEFT JOIN TownCity d" & _
                          " LEFT JOIN Province e" & _
                             " ON d.sProvIDxx = e.sProvIDxx" & _
                          " ON b.sTownIDxx = d.sTownIDxx" & _
                       " LEFT JOIN Occupation f" & _
                          " ON b.sOccptnID = f.sOccptnID" & _
                 " WHERE a.sClientID = b.sClientID" & _
                    " AND a.dTransact > '2012-01-01'" & _
                    " AND TIMESTAMPDIFF(MONTH, dTransact, SYSDATE()) > 1" & _
                    " AND a.cTranStat NOT IN ('3', '6', '7')" & _
                    " AND a.cPostedxx <> '6'" & _
                    " AND a.cPaymForm = '0'" & _
                    " AND (LENGTH(b.sMobileNo) = 11" & _
                       " OR LENGTH(b.sPhoneNox) = 11)" & _
                 " ORDER BY a.sClientID LIMIT 100000"

        p_oMaster = New DataTable
        Try
            p_oMaster = New DataTable
            p_oMaster = p_oAppDrvr.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
        End Try

        Return p_oMaster.Rows.Count > 0

endProc:
        Exit Function
errProc:
        ShowError(lsProcName & "( " & " ) ")
    End Function

    Private Function updateMPTLMClass(nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String = ""
        Dim lnRow As Integer
        Dim lnTotal As Long = 0
        Dim lnAmount As Long = 0
        Dim lsClassify As String

        lsProcName = "updateMPTLMClass"
        'On Error GoTo errProc

        'hard code ko muna ung classification
        '   lnTotal = p_oMaster("nCashAmtx") + IFNull(p_oMaster("xCCardAmt"), 0#)
        '   lsClassx = getClass(lnTotal, IFNull(p_oMaster("xCCardAmt"), 0#) > 0#, lnAmount)
        '   If lsClassx = "" Then GoTo endProc

        Select Case p_oMaster.Rows(nRow).Item("nCashAmtx")
            Case Is >= 20000
                lsClassify = "0001"
            Case Is >= 10000
                If Not IsDBNull(p_oMaster.Rows(nRow).Item("xCCardAmt")) Then
                    lsClassify = "0001"
                Else
                    lsClassify = "0002"
                End If
            Case Is >= 3000
                If Not IsDBNull(p_oMaster.Rows(nRow).Item("xCCardAmt")) Then
                    lsClassify = "0002"
                Else
                    lsClassify = "0003"
                End If
            Case Else
                lsClassify = "0004"
        End Select

        If p_oMaster.Rows(nRow).Item("sClassIDx") <> lsClassify Then
            If lnTotal > lnAmount Then
                If DateDiff("m", p_oMaster.Rows(nRow).Item("dLastCall"), p_oAppDrvr.SysDate) > 0 Then
                    lsSQL = "  sClassIDx = " & strParm(lsClassify) & _
                            ", dLastCall = NULL" & _
                            ", sRemarksx = ''"
                End If
            End If
        End If

        If lsSQL = "" Then
            updateMPTLMClass = updateCPSales(nRow)
            GoTo endProc
        End If

        lsSQL = " UPDATE TLM_Client SET" & _
                    lsSQL & _
                 " WHERE sClientID = " & strParm(p_oMaster.Rows(nRow).Item("sClientID"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        Return updateCPSales(nRow)

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function updateMCTLMClass(ByVal bCash As Boolean, nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String = ""
        Dim lnRow As Integer

        lsProcName = "updateMCTLMClass"
        'On Error GoTo errProc

        If IFNull(p_oMaster.Rows(nRow).Item("cSourceCd"), "") <> "MC" Then
            If Not IsDBNull(p_oMaster.Rows(nRow).Item("dLastCall")) Then
                If DateDiff("m", p_oMaster.Rows(nRow).Item("dLastCall"), p_oAppDrvr.SysDate) > 0 Then
                    lsSQL = "  cSourceCd = 'MC'" & _
                            ", sClassIDx = '0001'" & _
                            ", dLastCall = NULL" & _
                            ", sRemarksx = ''"
                End If
            End If
        End If

        If lsSQL = "" Then
            If Not bCash Then
                updateMCTLMClass = updateARMaster(nRow)
            Else
                updateMCTLMClass = updateSOMaster(nRow)
            End If
            GoTo endProc
        End If

        lsSQL = " UPDATE TLM_Client SET" & _
                    lsSQL & _
                 " WHERE sClientID = " & strParm(p_oMaster.Rows(nRow).Item("sClientID"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        If Not bCash Then
            Return updateARMaster(nRow)
        Else
            Return updateSOMaster(nRow)
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function addMPTLM(nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Long
        Dim lsClassify As String

        lsProcName = "addMPTLM"
        'On Error GoTo errProc

        'hard code ko muna ung classification
        '   lnTotal = (p_oMaster("nCashAmtx") + IFNull(p_oMaster("xCCardAmt"), 0#))
        '   lsClassify = getClass(lnTotal, IFNull(p_oMaster("xCCardAmt"), 0#) > 0#, lnAmount)
        '
        '   If lsClassify = "" Then GoTo endProc

        Select Case p_oMaster.Rows(nRow).Item("nCashAmtx")
            Case Is >= 20000
                lsClassify = "0001"
            Case Is >= 10000
                If Not IsDBNull(p_oMaster.Rows(nRow).Item("xCCardAmt")) Then
                    lsClassify = "0001"
                Else
                    lsClassify = "0002"
                End If
            Case Is >= 3000
                If Not IsDBNull(p_oMaster.Rows(nRow).Item("xCCardAmt")) Then
                    lsClassify = "0002"
                Else
                    lsClassify = "0003"
                End If
            Case Else
                lsClassify = "0004"
        End Select

        lsSQL = "INSERT INTO TLM_Client SET" & _
                    "  sClientID = " & strParm(p_oMaster.Rows(nRow).Item("sClientID")) & _
                    ", sClassIDx = " & strParm(lsClassify) & _
                    ", cSourceCd = 'MP'" & _
                    ", dBirthDte = " & dateParm(IFNull(p_oMaster.Rows(nRow).Item("dBirthDte"), "1900-01-01"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        Return updateCPSales(nRow)

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function addMCTLM(ByVal bCash As Boolean, nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Long

        lsProcName = "addMCTLM"
        'On Error GoTo errProc

        lsSQL = "INSERT IGNORE INTO TLM_Client SET" & _
                    "  sClientID = " & strParm(p_oMaster.Rows(nRow).Item("sClientID")) & _
                    ", sClassIDx = " & strParm("0001") & _
                    ", cSourceCd = 'MC'" & _
                    ", dBirthDte = " & dateParm(p_oMaster.Rows(nRow).Item("dBirthDte"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow < 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        If Not bCash Then
            Return updateARMaster(nRow)
        Else
            Return updateSOMaster(nRow)
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function updateCPSales(nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Integer

        lsProcName = "updateCPSales"
        'On Error GoTo errProc

        lsSQL = " UPDATE CP_SO_Master SET" & _
                    " cTranStat = " & strParm(Val(p_oMaster.Rows(nRow).Item("cTranStat") Xor 4)) & _
                 " WHERE sTransNox = " & strParm(p_oMaster.Rows(nRow).Item("sTransNox"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        Return True

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function updateARMaster(nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Integer

        lsProcName = "updateARMaster"
        'On Error GoTo errProc

        lsSQL = " UPDATE MC_AR_Master SET" & _
                    " cPostedxx = " & strParm(xeLogical.YES) & _
                 " WHERE sAcctNmbr = " & strParm(p_oMaster.Rows(nRow).Item("sAcctNmbr"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        Return True

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function updateSOMaster(nRow As Integer) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Integer

        lsProcName = "updateSOMaster"
        'On Error GoTo errProc

        lsSQL = " UPDATE MC_SO_Master SET" & _
                    " cPostedxx = '6'" & _
                 " WHERE sTransNox = " & strParm(p_oMaster.Rows(nRow).Item("sTransNox"))

        lnRow = p_oAppDrvr.ExecuteActionQuery(lsSQL)
        If lnRow <= 0 Then
            MsgBox("Unable to Update SO Master Status!", vbCritical, "Warning")
            Return False
            GoTo endProc
        End If

        Return True

endProc:
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Function getClass(ByVal Value As Long, _
                               ByVal CreditCard As Boolean, _
                               ByRef Amount As Long) As String
        Dim lsProcName As String
        Dim lrs As DataTable
        Dim lsSQL As String

        lsProcName = "getClass"
        'On Error GoTo errProc

        lsSQL = "SELECT *" & _
                 " FROM TLM_Class" & _
                 " WHERE nAmountxx <= " & CDbl(Value) & _
                    IIf(CreditCard = True, " OR cCredtCrd = '1'", "") & _
                 " ORDER BY nAmountxx DESC" & _
                 " LIMIT 1"

        lrs = New DataTable
        Try
            lrs = New DataTable
            lrs = p_oAppDrvr.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
        End Try

        If lrs.Rows.Count = 0 Then
            Return ""
            GoTo endProc
        End If

        Amount = lrs.Rows(0).Item("nAmountxx")
        Return lrs.Rows(0).Item("sClassIDx")

endProc:
        lrs = Nothing
        Exit Function
errProc:
        ShowError(lsProcName)
    End Function

    Private Sub ShowError(ByVal lsProcName As String)
        With p_oAppDrvr
            .ErrorLog(Err.Number & vbCrLf & Err.Description & vbCrLf & pxeMODULENAME & vbCrLf & lsProcName & vbCrLf & Erl())
        End With
    End Sub
End Class
