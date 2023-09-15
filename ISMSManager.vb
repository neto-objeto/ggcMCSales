'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Incoming SMS Manager
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
'  Kalyptus [ 04/29/2016 02:20 pm ]
'      Started creating this object.
'  Kalyptus [ 04/30/2016 11:20 pm ]
'       Added a parent property
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class ISMSManager
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_sParent As String
    Private p_bCallNext As Boolean
    Private p_bUpdtSMS As Boolean
    Private p_sCallDate As String

    Private Const p_sMasTable As String = "SMS_Incoming"
    Private Const p_sMsgHeadr As String = "Incoming SMS Message"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)


    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTMstr(0).Item(Index) = value
            End If
        End Set

    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTMstr(0).Item(Index) = value
            End If
        End Set
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public Property isCallAgain() As Boolean
        Get
            Return p_bCallNext
        End Get
        Set(ByVal value As Boolean)
            p_bCallNext = value
        End Set
    End Property

    Public Property isSMSUpdate() As Boolean
        Get
            Return p_bUpdtSMS
        End Get
        Set(ByVal value As Boolean)
            p_bUpdtSMS = value
        End Set
    End Property

    Public Property CallDate() As String
        Get
            Return p_sCallDate
        End Get
        Set(ByVal value As String)
            If IsDate(value) Then
                p_sCallDate = value
            End If
        End Set
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    'Public Function NewTransaction()
    Public Function NewTransaction() As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())
        Call initMaster()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    'Public Function OpenTransaction(String)
    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    'Public Function SaveTransaction
    'This object does not implement Update
    Public Function SaveTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        'Save master table 
        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)
        Else
            'Create the statement
            If p_bCallNext And IsDate(p_sCallDate) Then
                'All cTranstat of rescheduled calls should be Open 
                lsSQL = "UPDATE " & p_sMasTable & _
                       " SET dFollowUp = " & datetimeParm(p_sCallDate) & _
                        IIf(p_bUpdtSMS = True, ", sMessagex = " & strParm(p_oDTMstr(0).Item("sMessagex")), "") & _
                          ", cTranStat = '0'" & _
                       " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            Else
                If p_bUpdtSMS Then
                    'update the message
                    lsSQL = "UPDATE " & p_sMasTable & _
                           " SET   sMessagex = " & strParm(p_oDTMstr(0).Item("sMessagex")) & _
                                ", sModified = " & strParm(p_oApp.UserID) & _
                           " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                Else
                    'Set the cTranStat of this call as the transtat of the source transaction
                    lsSQL = "UPDATE " & p_sMasTable & _
                           " SET cTranStat = '3'" & _
                           " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                End If
            End If
            Call p_oApp.Execute(lsSQL, p_sMasTable)

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, Format(p_oApp.SysDate, ggcAppDriver.xsDATE_TIME))
            End If

            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, p_sMasTable)
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
    End Function

    Public Function Get2Read() As String
        Dim lsSQL As String
        Dim loDta As DataTable
        lsSQL = "SELECT sTransNox" & _
                " FROM " & p_sMasTable & _
                " WHERE IFNULL(cTranStat, '0') = '0'" & _
                    " AND IFNULL(dFollowUp, '1900-01-01') = " & dateParm(ggcAppDriver.xsNULL_DATE) & _
                    " AND sMessagex NOT LIKE 'REG %'" & _
                    " AND sMessagex NOT LIKE 'CODEAPPR %'" & _
                " ORDER BY dTransact ASC" & _
                " LIMIT 1"

        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count <= 0 Then
            lsSQL = ""
        Else
            lsSQL = loDta(0).Item("sTransNox")
        End If

        Return lsSQL
    End Function

    Public Function Schedule() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        p_oApp.BeginTransaction()

        Dim lsSQL As String

        p_oDTMstr(0).Item("cTranStat") = "1"

        lsSQL = "UPDATE " & p_sMasTable & _
               " SET cTranStat = '1'" & _
               " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
        p_oApp.Execute(lsSQL, p_sMasTable)

        p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function Close() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_sParent = "" Then p_oApp.BeginTransaction()

        Dim lsSQL As String

        p_oDTMstr(0).Item("cTranStat") = "2"

        lsSQL = "UPDATE " & p_sMasTable & _
               " SET cTranStat = '2'" & _
               " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
        p_oApp.Execute(lsSQL, p_sMasTable)

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function Disregard() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_sParent = "" Then p_oApp.BeginTransaction()

        Dim lsSQL As String

        p_oDTMstr(0).Item("cTranStat") = "3"

        lsSQL = "UPDATE " & p_sMasTable & _
               " SET cTranStat = '3'" & _
               " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
        p_oApp.Execute(lsSQL, p_sMasTable)

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dfollowup", "dreadxxxx"
                    p_oDTMstr(0).Item(lnCtr) = ggcAppDriver.xsNULL_DATE
                Case "nnoretryx"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case "dmodified", "smodified"
                Case "ctranstat", "creadxxxx", "csubscrbr"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT a.sTransnox" & _
                    ", a.dTransact" & _
                    ", a.sSourceCd" & _
                    ", a.sMessagex" & _
                    ", a.sMobileNo" & _
                    ", a.cSubscrbr" & _
                    ", a.dFollowUp" & _
                    ", a.nNoRetryx" & _
                    ", a.cReadxxxx" & _
                    ", a.dReadxxxx" & _
                    ", a.cTranStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.dTransact" & _
                    ", a.sMobileNo" & _
                    ", a.dTransact" & _
                    ", a.sSourceCd" & _
              " FROM " & p_sMasTable & " a"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub
End Class


