'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC Product Inquiry Object
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
'  Kalyptus [ 04/02/2016 10:50 am ]
'      Started creating this object.
'  Kalyptus [ 04/30/2016 11:20 pm ]
'       Added a parent property
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ggcAppDriver
Imports ggcClient
Imports System.IO
Imports ADODB

Public Class MCProductInquiry
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String

    Private p_oClient As ggcClient.Client

    Private Const p_sMasTable As String = "MC_Product_Inquiry"
    Private Const p_sMsgHeadr As String = "MC Product Inquiry"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sClientNm
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sAddressx) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case 82 ' sModelNme
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sModelNme) = "" Then
                            getModel(3, 82, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sModelNme
                    Case 83 ' sColorNme
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sColorNme) = "" Then
                            getColor(4, 83, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sColorNme
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
                    Case 80 ' sClientNm
                        getClient(2, 80, value, False, False)
                    Case 81 ' sAddressx
                    Case 82 ' sModelNme
                        getModel(3, 82, value, False, False)
                    Case 83 ' sColorNme
                        getColor(4, 83, value, False, False)
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
                    Case "sclientnm" ' 80 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" ' 81 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sAddressx) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case "smodelnme" ' 82 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sModelNme) = "" Then
                            getModel(3, 82, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sModelNme
                    Case "scolornme" ' 83 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sColorNme) = "" Then
                            getColor(4, 83, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sColorNme
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
                        getClient(2, 80, value, False, False)
                    Case "saddressx"
                    Case "smodelnme"
                        getModel(3, 82, value, False, False)
                    Case "scolornme"
                        getColor(4, 83, value, False, False)
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    'Property ()
    Public ReadOnly Property BranchCode() As String
        Get
            Return p_sBranchCd
        End Get
    End Property

    Public ReadOnly Property BranchName() As String
        Get
            Return p_sBranchNm
        End Get
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
        Call InitOthers()

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

        Call InitOthers()

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    'Public Function SearchWithCondition(String)
    Public Function SearchWithCondition(ByVal fsFilter As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Browse, fsFilter)
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        ElseIf p_oDTMstr.Rows.Count = 1 Then
            Return OpenTransaction(p_oDTMstr(0).Item("sTransNox"))
        Else
            'KwikBrowse here!
            Return True
        End If
    End Function

    'Public Function SearchTransaction(String, Boolean, Boolean=False)
    Public Function SearchTransaction( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oOthersx.sClientNm Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = getSQ_Browse()
        End If

        If p_sBranchCd <> "" Then
            lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
        End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm(p_oApp.BranchCode & "%" & fsValue)
        Else
            lsFilter = "a.sCompnyNm like " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , "Trans No»Client»Date", _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
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

        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        If Not p_oClient.SaveClient Then
            MsgBox("Unable to save client info!", vbOKOnly, p_sMsgHeadr)
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            Return False
        End If

        p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")

        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            'Save master table 

            'save user that creates the inquiry
            p_oDTMstr(0).Item("sCreatedx") = p_oApp.UserID
            p_oDTMstr(0).Item("dCreatedx") = p_oApp.getSysDate

            p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)

            p_oApp.Execute(lsSQL, p_sMasTable)
        End If

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    'Public Function CancelTransaction
    Public Function CancelTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oApp.UserLevel And (xeUserRights.ENGINEER + xeUserRights.MANAGER) = 0 Then
            MsgBox("This user is not allowed to cancel this transaction!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to cancel!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already voided and is no longer allowed to void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "3"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    'Public Function VoidTransaction
    Public Function VoidTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oApp.UserLevel And (xeUserRights.ENGINEER + xeUserRights.MANAGER) = 0 Then
            MsgBox("This user is not allowed to void this transaction!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already void!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "4"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    'Public Function CloseTransaction()
    Public Function CloseTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Request was already acknowledged and is no longer allowed to re-acknowledge!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already voided!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        If p_oDTMstr(0).Item("cTranStat") = "0" Then
            p_oDTMstr(0).Item("cTranStat") = "1"

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update MC_Product_Inquiry status.", MsgBoxStyle.Information, "Warning")
                GoTo endWithRoll
            End If

            Dim lsMobile As String = ""
            Dim loClient As ggcClient.Client
            loClient = New ggcClient.Client(p_oApp)
            loClient.Parent = "MCProductInquiry"

            If loClient.OpenClient(p_oDTMstr(0).Item("sClientID")) Then lsMobile = loClient.Master("sMobileNo")

            If lsMobile = "" Then
                MsgBox("Customer has no mobile number.", MsgBoxStyle.Information, "Warning")
                GoTo endWithRoll
            End If

            lsSQL = p_oApp.getConfiguration("WebSvr")
            lsSQL = httpsGET(lsSQL & "telemarketing/getNetwork.php?mobile=" & lsMobile)

            If lsSQL = "" Then
                MsgBox("Unable to get network classification." & vbCrLf & vbCrLf & _
                       "Please inform MIS Department.", MsgBoxStyle.Information, "Warning")
                GoTo endWithRoll
            ElseIf Len(lsSQL) > 1 Then
                MsgBox("Error getting network classification." & vbCrLf & vbCrLf & _
                       "Please inform MIS Department.", MsgBoxStyle.Information, "Warning")
                GoTo endWithRoll
            End If

            'convert the inquiry to leads
            lsSQL = "INSERT INTO Call_Outgoing SET " & _
                    "  sTransNox = " & strParm(GetNextCode("Call_Outgoing", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                    ", dTransact = " & dateParm(p_oApp.SysDate) & _
                    ", sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) & _
                    ", sMobileNo = " & strParm(lsMobile) & _
                    ", sRemarksx = ''" & _
                    ", sReferNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                    ", sSourceCD = " & strParm("INQR") & _
                    ", cTranStat = '1'" & _
                    ", sAgentIDx = " & strParm(p_oDTMstr(0).Item("sCreatedx")) & _
                    ", nNoRetryx = 0" & _
                    ", cSubscrbr = " & strParm(lsSQL) & _
                    ", cCallStat = '0'" & _
                    ", cTLMStatx = '0'" & _
                    ", cSMSStatx = '0'" & _
                    ", nSMSSentx = 0" & _
                    ", sModified = " & strParm(p_oApp.UserID) & _
                    ", dModified = " & strParm(p_oApp.SysDate)

            If p_oApp.Execute(lsSQL, "Call_Outgoing", Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to add inquiry to TLM Leads.", MsgBoxStyle.Information, "Warning")
                GoTo endWithRoll
            End If
        End If

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
endWithRoll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()
        Return False
    End Function

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 80 ' sClientNm
                getClient(2, 80, fsValue, False, True)
            Case 82 ' sModelNme
                getModel(3, 82, fsValue, False, True)
            Case 83 ' sColorNme
                getColor(4, 83, fsValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dmodified", "smodified", "dtargetxx", "dfollowup", "dcreatedx"
                Case "ctranstat", "cpurctype"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
        p_oOthersx.sModelNme = ""
        p_oOthersx.sColorNme = ""
    End Sub

    Private Function isEntryOk() As Boolean
        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Inquiry was posted! Posted application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Inquiry was cancelled! Cancelled application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("sClientID") = "" Then
            MsgBox("Invalid Client info detected!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("sModelIDx") = "" Then
            MsgBox("Invalid model detected!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sClientNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sClientNm And fsValue <> "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oApp)
        loClient.Parent = "MCProductInquiry"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oClient = loClient
                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oOthersx.sClientNm = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")

                p_oOthersx.sAddressx = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            Else
                p_oDTMstr(0).Item("sClientID") = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sAddressx = ""
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                p_oClient = loClient
                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oOthersx.sClientNm = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")
                p_oOthersx.sAddressx = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            End If
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
        RaiseEvent MasterRetrieved(81, p_oOthersx.sAddressx)
    End Sub

    ''This method implements a search master where id and desc are not joined.
    'Private Sub getClient(ByVal fnColIdx As Integer _
    '                    , ByVal fnColDsc As Integer _
    '                    , ByVal fsValue As String _
    '                    , ByVal fbIsCode As Boolean _
    '                    , ByVal fbIsSrch As Boolean)

    '    'Compare the value to be search against the value in our column
    '    If fbIsCode Then
    '        If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sClientNm <> "" Then Exit Sub
    '    Else
    '        If fsValue = p_oOthersx.sClientNm And fsValue <> "" Then Exit Sub
    '    End If



    '    Dim lsSQL As String
    '    lsSQL = "SELECT" & _
    '                   "  a.sClientID" & _
    '                   ", a.sCompnyNm sClientNm" & _
    '                   ", CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" & _
    '           " FROM Client_Master a" & _
    '                  " LEFT JOIN TownCity b ON a.sTownIDxx = b.sTownIDxx" & _
    '                  " LEFT JOIN Province c ON b.sProvIDxx = c.sProvIDxx" & _
    '           " WHERE a.cRecdStat = '1'"

    '    'Are we using like comparison or equality comparison
    '    If fbIsSrch Then
    '        Dim loRow As DataRow = KwikSearch(p_oApp _
    '                                         , lsSQL _
    '                                         , True _
    '                                         , fsValue _
    '                                         , "sClientID»sClientNm»xAddressx" _
    '                                         , "ID»Client Name»Address", _
    '                                         , "a.sClientID»a.sCompnyNm»CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode)" _
    '                                         , IIf(fbIsCode, 0, 1))
    '        If IsNothing(loRow) Then
    '            p_oDTMstr(0).Item(fnColIdx) = ""
    '            p_oOthersx.sClientNm = ""
    '            p_oOthersx.sAddressx = ""
    '        Else
    '            p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sClientID")
    '            p_oOthersx.sClientNm = loRow.Item("sClientNm")
    '            p_oOthersx.sAddressx = loRow.Item("xAddressx")
    '        End If

    '        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
    '        Exit Sub

    '    End If

    '    If fsValue = "" Then
    '        lsSQL = AddCondition(lsSQL, "0=1")
    '    Else
    '        If fbIsCode Then
    '            lsSQL = AddCondition(lsSQL, "a.sClientID = " & strParm(fsValue))
    '        Else
    '            lsSQL = AddCondition(lsSQL, "a.sCompnyNm = " & strParm(fsValue))
    '        End If
    '    End If

    '    Dim loDta As DataTable
    '    loDta = p_oApp.ExecuteQuery(lsSQL)

    '    If loDta.Rows.Count = 0 Then
    '        p_oDTMstr(0).Item(fnColIdx) = ""
    '        p_oOthersx.sClientNm = ""
    '        p_oOthersx.sAddressx = ""
    '    ElseIf loDta.Rows.Count = 1 Then
    '        p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sClientID")
    '        p_oOthersx.sClientNm = loDta(0).Item("sClientNm")
    '        p_oOthersx.sAddressx = loDta(0).Item("xAddressx")
    '    End If

    '    RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)

    'End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getModel(ByVal fnColIdx As Integer _
                       , ByVal fnColDsc As Integer _
                       , ByVal fsValue As String _
                       , ByVal fbIsCode As Boolean _
                       , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sModelNme <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sModelNme And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sModelIDx" & _
                       ", a.sModelNme" & _
               " FROM MC_Model a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sModelIDx»sModelNme" _
                                             , "ID»Model", _
                                             , "a.sModelIDx»a.sModelNme" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sModelNme = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sModelIDx")
                p_oOthersx.sModelNme = loRow.Item("sModelNme")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sModelNme)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sModelIDx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sModelNme = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sModelNme = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sModelIDx")
            p_oOthersx.sModelNme = loDta(0).Item("sModelNme")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sModelNme)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getColor(ByVal fnColIdx As Integer _
                       , ByVal fnColDsc As Integer _
                       , ByVal fsValue As String _
                       , ByVal fbIsCode As Boolean _
                       , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sColorNme <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sColorNme And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sColorIDx" & _
                       ", a.sColorNme" & _
               " FROM Color a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sColorIDx»sColorNme" _
                                             , "ID»Model", _
                                             , "a.sColorIDx»a.sColorNme" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sColorNme = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sColorIDx")
                p_oOthersx.sColorNme = loRow.Item("sColorNme")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sColorNme)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "0=1")
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sColorIDx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sColorNme = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sColorNme = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sColorIDx")
            p_oOthersx.sColorNme = loDta(0).Item("sColorNme")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sColorNme)

    End Sub

    'This method implements a search master where id and desc are not joined.
    Public Sub searchBranch(ByVal fsValue As String _
                           , ByVal fbIsCode As Boolean _
                           , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_sBranchCd And fsValue <> "" Then Exit Sub
        Else
            If fsValue = p_sBranchNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        Dim lsFilter As String
        lsSQL = "SELECT" & _
                       "  a.sBranchCD" & _
                       ", a.sBranchNm" & _
               " FROM Branch a" & _
               " WHERE a.cRecdStat = '1'"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp, lsSQL, True, fsValue, "sBranchCD»sBranchNm", "Code»Branch", "", "a.sBranchCD»a.sBranchNm", IIf(fbIsCode, 0, 1))

            If Not IsNothing(loRow) Then
                p_sBranchCd = loRow.Item("sBranchCD")
                p_sBranchNm = loRow.Item("sBranchNm")
            End If
            Exit Sub

        End If

        If fsValue = "" Then
            lsFilter = "0=1"
        Else
            If fbIsCode Then
                lsFilter = "a.sBranchCD = " & strParm(fsValue)
            Else
                lsFilter = "a.sBranchNm = " & strParm(fsValue)
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(AddCondition(lsSQL, lsFilter))

        If loDta.Rows.Count = 0 Then
            p_sBranchCd = ""
            p_sBranchNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_sBranchCd = loDta(0).Item("sBranchCD")
            p_sBranchNm = loDta(0).Item("sBranchNm")
        End If
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT a.sTransnox" & _
                    ", a.dTransact" & _
                    ", a.sClientID" & _
                    ", a.sModelIDx" & _
                    ", a.sColorIDx" & _
                    ", a.sInquiryX" & _
                    ", a.dTargetxx" & _
                    ", a.dFollowUp" & _
                    ", a.cPurcType" & _
                    ", a.sRemarks1" & _
                    ", a.sRemarks2" & _
                    ", a.sSourceNo" & _
                    ", a.cTranStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                    ", a.sCreatedx" & _
                    ", a.dCreatedx" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
              " WHERE a.sClientID = b.sClientID"
    End Function

    Private Function httpsGET(ByVal fsURL As String) As String
        Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri(fsURL))

        request.Method = System.Net.WebRequestMethods.Http.Get

        Dim response As System.Net.HttpWebResponse = request.GetResponse()

        Dim dataStream As System.IO.Stream = response.GetResponseStream()
        Dim reader As System.IO.StreamReader = New System.IO.StreamReader(dataStream)
        Dim lsValue As String = reader.ReadToEnd()
        reader.Close()
        response.Close()

        Return lsValue
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "MCProductInquiry"
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sAddressx As String
        Public sModelNme As String
        Public sColorNme As String
    End Class
End Class
