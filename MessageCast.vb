'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Message Cast Object
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
'  iMac [ 08/10/2017 05:00 pm ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports ggcAppDriver

Public Class MessageCast
    Private Const pxeMasterTable As String = "Text_Mktg_Master"
    Private Const pxeDetailTable As String = "Text_Mktg_Detail"
    Private Const pxeMessageHead As String = "MessageCast"

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oOthersx As New Others
    Private p_nEditMode As xeEditMode
    Private p_nTranStat As xeTranStat
    Private p_sParent As String
    Private p_sBranchCd As String

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    WriteOnly Property Parent
        Set(value)
            p_sParent = value
        End Set
    End Property

    WriteOnly Property Branch
        Set(value)
            p_sBranchCd = value
        End Set
    End Property

    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sCompnyNm
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sCompnyNm) = "" Then
                            getCompany(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sCompnyNm
                        getCompany(2, 80, value, False, False)
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set

    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "scompnynm" ' 80 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sCompnyNm) = "" Then
                            getCompany(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "sclientnm"
                        getCompany(2, 80, value, False, False)
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 80 ' sClientNm
                getCompany(2, 80, fsValue, False, True)
        End Select
    End Sub

    Public Function NewTransaction() As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQL_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())

        Call initMaster()
        Call initOthers()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Public Function SaveTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            'Save master table 

            p_oDTMstr(0).Item("sTransNox") = GetNextCode(pxeMasterTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
            lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTable, , p_oApp.UserID, p_oApp.SysDate)

            p_oApp.Execute(lsSQL, pxeMasterTable)
        End If

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function CancelTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oApp.UserLevel And (xeUserRights.ENGINEER + xeUserRights.MANAGER) = 0 Then
            MsgBox("This user is not allowed to cancel this transaction!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to cancel!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already voided and is no longer allowed to void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        Dim lsSQL As String

        p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "3"
        lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, pxeMasterTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function VoidTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oApp.UserLevel And (xeUserRights.ENGINEER + xeUserRights.MANAGER) = 0 Then
            MsgBox("This user is not allowed to void this transaction!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        Dim lsSQL As String

        p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "4"
        lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, pxeMasterTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function CloseTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oApp.UserLevel And (xeUserRights.ENGINEER + xeUserRights.MANAGER) = 0 Then
            MsgBox("This user is not allowed to close this transaction!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to re-acknowledge!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already voided!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, pxeMessageHead)
            Return False
        End If

        Dim lsSQL As String

        p_oApp.BeginTransaction()

        If p_oDTMstr(0).Item("cTranStat") = "0" Then
            p_oDTMstr(0).Item("cTranStat") = "1"

            lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
            p_oApp.Execute(lsSQL, pxeMasterTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))
        End If

        p_oApp.CommitTransaction()

        Return True
    End Function

    Private Sub getCompany(ByVal fnColIdx As Integer _
                           , ByVal fnColDsc As Integer _
                           , ByVal fsValue As String _
                           , ByVal fbIsCode As Boolean _
                           , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sCompnyNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sCompnyNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" +
                    "  a.sCompnyID" +
                    ", a.sCompnyNm" +
                    ", a.sCompnyCd" +
                    ", a.cRecdStat" +
               " FROM Company a" & _
               " WHERE a.cRecdStat = '1'"

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sCompnyID»sCompnyNm»sCompnyCd" _
                                             , "ID»Name»Code", _
                                             , "a.sCompnyID»a.sCompnyNm»a.sCompnyCd" _
                                             , IIf(fbIsCode, 0, 1))

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oDTMstr(0).Item(4) = ""
                p_oOthersx.sCompnyNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sCompnyID")
                p_oDTMstr(0).Item(4) = loRow.Item("sCompnyCd")
                p_oOthersx.sCompnyNm = loRow.Item("sCompnyNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCompnyNm)
            RaiseEvent MasterRetrieved(4, p_oDTMstr(0).Item(4))
            Exit Sub
        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sCompnyID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sCompnyNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oDTMstr(0).Item(4) = ""
            p_oOthersx.sCompnyNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sCompnyID")
            p_oDTMstr(0).Item(4) = loDta(0).Item("sCompnyCd")
            p_oOthersx.sCompnyNm = loDta(0).Item("sCompnyNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCompnyNm)
        RaiseEvent MasterRetrieved(4, p_oDTMstr(0).Item(4))
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(pxeMasterTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "dtransact", "dschedfrm", "dschedtru"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dmodified", "smodified"
                Case "ctranstat"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "nallocatn", "nsuccessx"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Function isEntryOk() As Boolean
        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Record was posted! Posted application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMessageHead)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Record was cancelled! Cancelled application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMessageHead)
            Return False
        End If

        If p_oDTMstr(0).Item("sCompnyID") = "" Then
            MsgBox("Invalid company detected!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMessageHead)
            Return False
        End If

        Return True
    End Function

    Private Function getSQL_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", b.sCompnyNm" & _
                    ", a.dTransact" & _
              " FROM " & pxeMasterTable & " a" & _
                    ", Company b" & _
              " WHERE a.sCompnyID = b.sCompnyID"
    End Function

    Private Function getSQL_Master() As String
        Return "SELECT" +
                     "  sTransNox" +
                     ", dTransact" +
                     ", sCompnyID" +
                     ", sMessagex" +
                     ", sTextCode" +
                     ", dSchedFrm" +
                     ", dSchedTru" +
                     ", nAllocatn" +
                     ", nSuccessx" +
                     ", sRepCPNox" +
                     ", sRepEmail" +
                     ", cTranStat" +
                     ", sModified" +
                     ", dModified" +
                " FROM " & pxeMasterTable
    End Function

    Public Function SearchTransaction( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sClientNm") Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQL_Browse, "a.cTranStat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = getSQL_Browse()
        End If

        If p_sBranchCd <> "" Then
            lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
        End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm(p_oApp.BranchCode & "%" & fsValue)
        Else
            lsFilter = "a.sClientNm like " & strParm("%" & fsValue)
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sCompnyNm»dTransact" _
                                        , "Trans No»Company»Date", _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If

    End Function

    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQL_Master, "sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        Call initOthers()

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    Private Sub initOthers()
        p_oOthersx.sCompnyNm = ""
    End Sub

    Private Class Others
        Public sCompnyNm As String
    End Class
End Class
