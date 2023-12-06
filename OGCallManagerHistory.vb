
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC AR Master Object
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
'  Kalyptus [ 06/04/2016 11:25 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient

Public Class OGCallManagerHistory
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetail As DataTable

    Private p_oClient As ggcClient.Client
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_oCallInfosx As New OGCallInfo
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sBranchCD As String
    Private p_bCallNext As Boolean
    Private p_sCallDate As String

    Public Const pxeCOS_TEXTMESSAGE As String = "ISMS" 'sms incoming
    Public Const pxeCOS_INCOMINGCALL As String = "CALL" 'call incoming
    Public Const pxeCOS_INQUIRY As String = "INQR" 'mc product inquiry
    Public Const pxeCOS_REFERRAL As String = "RFRL" 'referral
    Public Const pxeCOS_HOTLINE As String = "TXHL"
    Public Const pxeCOS_REQUEST As String = "RQST"
    Public Const pxeCOS_RANDOMCALL As String = "TLMC" 'random leads from mp customers/mc customers
    Public Const pxeCOS_LENDING As String = "LEND"  'lending/cash loan
    Public Const pxeCOS_BYAHENGFIESTA As String = "GBF"  'activity inqry byaheng fiesta
    Public Const pxeCOS_FREESERVICE As String = "FSCU"  'activity inqry fscu
    Public Const pxeCOS_DISPLAYCARAVAN As String = "DC" 'activity inqry display caravan
    Public Const pxeCOS_OTHERS As String = "OTH" 'activity inqry others
    Public Const pxeCOS_MCSALES As String = "MCSO" 'mp leads from mc customers
    Public Const pxeCOS_MPINQR = "MPIn" 'mp product inquiry     
    Public Const pxeCOS_CA = "MCCA" 'credit application
    Public Const pxeCOS_GANADO = "GNDO" 'Ganado Online


    Private Const p_sMasTable As String = "Call_Outgoing"
    Private Const p_sMsgHeadr As String = "TLM Sales"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)
    Public Event DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer, _
                              ByVal Value As Object)

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Branch As String
        Get
            Return p_sBranchCD
        End Get
        Set(ByVal value As String)
            'If Product ID is LR then do allow changing of Branch
            If p_oApp.ProductID = "TeleMktg" Then
                p_sBranchCD = value
            End If
        End Set
    End Property


    Public Property isCallAgain() As Boolean
        Get
            Return p_bCallNext
        End Get
        Set(ByVal value As Boolean)
            p_bCallNext = value
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

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 0 ' sClientNm
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" Then
                            getClient(2, 0, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 2 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" Then
                            getClient(2, 2, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                   
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
              
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case "sclientnm" ' 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" ' 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sAddressx) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
               
            End If
        End Set
    End Property
    Public Property Detail(ByVal Row As Integer, ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index

                    'Case 2 ' sClientNm
                    '    If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oCallInfosx.xClientNmeCallInfo) = "" Then
                    '        'getClients(2, 2, p_oDTMstr(0).Item(2), True, False)
                    '    End If
                    '    Return p_oCallInfosx.xClientNmeCallInfo
                    'Case 3 ' sMobileNo
                    '    Return p_oDTMstr(0).Item(2)

                    Case Else
                        Return p_oDTDetail(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 8
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = Format(CDate(value), "yyyy-MM-dd")

                        End If


                        RaiseEvent DetailRetrieved(Row, Index, p_oDTDetail(0).Item(Index))

                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    'Property Detail(String)
    Public Property Detail(ByVal Row As Integer, ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "sclientid" ' 1 
                        If Trim(IFNull(p_oDTDetail(Row).Item(1))) <> "" Then
                            getClientDetail(Row, 1, 1, p_oDTDetail(Row).Item(1), True, False)
                            Return p_oCallInfosx.xClientNmeCallInfo
                        End If
                    Case "sagentidx" ' 3 
                        p_oCallInfosx.xAgentNme = getAgent(p_oDTDetail(Row).Item(3))
                        Return p_oCallInfosx.xAgentNme

                    Case Else
                        Return p_oDTDetail(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)

                    Case "sremarksx" '9
                        If (value <> "") Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        
                        RaiseEvent DetailRetrieved(Row, Index, p_oDTDetail(0).Item(Index))
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

    Public Function GetItemCount() As Integer
        If p_oDTDetail Is Nothing Then Return 0

        Return p_oDTDetail.Rows.Count
    End Function

    'Public Function NewTransaction()
    Public Function NewTransaction() As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())
        Call initMaster()
        Call InitOthers()
        Call InitOGCallInfo()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    'Public Function OpenMasterList(String)
    Public Function OpenMasterList(ByVal fsClientID As String) As Boolean
        Dim lsMasterSQL As String
        Dim lsDetailSQL As String

        lsMasterSQL = AddCondition(getSQ_Master(), "a.sClientID = '" & fsClientID & "' LIMIT 1")

        p_oDTMstr = p_oApp.ExecuteQuery(lsMasterSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            GoTo endWithWarning
            Return False
        End If

        If p_oDTMstr.Columns.Contains("sMobileNo") Then
            lsDetailSQL = AddCondition(getSQ_Detail(), "a.sMobileNo = '" & p_oDTMstr.Rows(0)("sMobileNo").ToString() & "'")
            p_oDTDetail = p_oApp.ExecuteQuery(lsDetailSQL)

            For rowIndex As Integer = 0 To p_oDTDetail.Rows.Count - 1
                Dim detailRow As DataRow = p_oDTDetail.Rows(rowIndex)
                RaiseEvent DetailRetrieved(rowIndex, -1, detailRow)
            Next

            If p_oDTDetail.Rows.Count > 0 Then
                p_nEditMode = xeEditMode.MODE_READY

                Return True
            Else
                Return False
            End If
        End If
endWithWarning:
        MsgBox("No Leads Found for this Transaction!" & _
                 vbCrLf & " Can Not Process Transaction", vbCritical, "Warning")
        GoTo endProc
endProc:
        p_oDTMstr = Nothing
        p_oDTDetail = Nothing
        Exit Function
    End Function

    'Public Function SearchTransaction(String, Boolean, Boolean=False)
    Public Function SearchTransaction(
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        'If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
        '    If fbByCode Then
        '        If fsValue = p_oDTMstr(0).Item("sMobileNo") Then Return True
        '    Else
        '        If fsValue = p_oCallInfosx.xClientNmeCallInfo Then Return True
        '    End If
        'End If

        'Initialize SQL filter
        'If p_nTranStat >= 0 Then
        '    lsSQL = AddCondition(getSQ_Browse, "ORDER BY b.sCompnyNm")
        'Else
        lsSQL = getSQ_Browse()
        'End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sMobileNo LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "b.sCompnyNm like " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sMobileNo»sCompnyNm" _
                                        , "Mobile No»Client Name",
                                        , "a.sMobileNo»b.sCompnyNm" _
                                        , IIf(fbByCode, 0, 1))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            'compare old client id of old 
            If Not IsNothing(p_oDTMstr) Then
                If Not p_oDTMstr(0).Item(2) = loDta.Item("sClientID") Then
                    p_oOthersx.sClientNm = ""
                End If

            End If
            Return OpenMasterList(loDta.Item("sClientID"))
            End If
    End Function

    'Public Function SaveTransaction
    'This object does not implement Update
    Public Function SaveTransaction(ByVal lnRow As Integer) As Boolean
        Dim lsSQL As String
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
        p_nEditMode = xeEditMode.MODE_READY Or _
        p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If Not IsNothing(p_oDTDetail) Then
            p_oApp.BeginTransaction()
        End If
        Try
            'UPDATE THE SOURCE TRANSACTION OF THE OUTGOING CALL
            'Determine the table
            Dim lsTableNme As String
            Select Case UCase(p_oDTDetail(lnRow).Item("sSourceCD"))
                Case pxeCOS_TEXTMESSAGE
                    lsTableNme = "SMS_Incoming"
                Case pxeCOS_INQUIRY
                    lsTableNme = "MC_Product_Inquiry"
                Case pxeCOS_REFERRAL
                    lsTableNme = "MC_Referral"
                Case pxeCOS_RANDOMCALL
                    lsTableNme = "TLM_Client"
                Case pxeCOS_LENDING
                    lsTableNme = "LR_Pre_Approve"
                Case pxeCOS_BYAHENGFIESTA, pxeCOS_DISPLAYCARAVAN, pxeCOS_FREESERVICE, pxeCOS_OTHERS
                    lsTableNme = "Activity_Inquiry"
                Case pxeCOS_MPINQR
                    lsTableNme = "MP_Product_Inquiry"
                Case pxeCOS_CA
                    lsTableNme = "MC_Credit_Application"
                Case pxeCOS_GANADO
                    lsTableNme = "Ganado_Online"
                Case Else 'pxeCOS_INCOMINGCALL
                    lsTableNme = "Call_Incoming"
            End Select

            'Create the statement
            If p_bCallNext And IsDate(p_sCallDate) Then
                'All cTranstat of rescheduled calls should be Open 
                Select Case LCase(lsTableNme)
                    Case "mc_product_inquiry", "mp_product_inquiry"
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                    ", cTranStat = '0'" &
                                    ", sCreatedx = " & strParm(p_oApp.UserID) &
                                " WHERE sTransNox = " & strParm(p_oDTDetail(lnRow)("sReferNox"))
                    Case "tlm_client"
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                    ", cTranStat = '0'" &
                                    ", sAgentIDx = " & strParm(p_oApp.UserID) &
                                " WHERE sClientID = " & strParm(p_oDTDetail(lnRow).Item("sClientID"))
                    Case "activity_inquiry"
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                    ", cTranStat = '0'" &
                                    ", sAgentIDx = " & strParm(p_oApp.UserID) &
                                " WHERE sInqryIDx = " & strParm(p_oDTDetail(lnRow).Item("sReferNox"))
                    Case "mc_credit_application"
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  cTLMStatx = '0'" &
                                    ", sTLMAgent = " & strParm(p_oApp.UserID) &
                                    ", dFollowUp = " & datetimeParm(p_sCallDate) &
                                " WHERE sTransNox = " & strParm(p_oDTDetail(lnRow).Item("sReferNox"))
                    Case "ganado_online"
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  cTranStat = '1'" &
                                    ", sTLMAgent = " & strParm(p_oApp.UserID) &
                                    ", dFollowUp = " & datetimeParm(p_sCallDate) &
                                " WHERE sTransNox = " & strParm(p_oDTDetail(lnRow).Item("sReferNox"))
                    Case Else
                        lsSQL = "UPDATE " & lsTableNme &
                                " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                    ", cTranStat = '0'" &
                                    ", sAgentIDx = " & strParm(p_oApp.UserID) &
                               " WHERE sTransNox = " & strParm(p_oDTDetail(lnRow).Item("sTransNox"))
                End Select


                Call p_oApp.Execute(lsSQL, lsTableNme)

                p_oDTDetail(lnRow)("dModified") = p_oApp.SysDate
                lsSQL = ADO2SQL(p_oDTDetail, p_sMasTable, "sTransNox = " & strParm(p_oDTDetail(lnRow).Item("sTransNox")), p_oApp.UserID)
            End If

            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, p_sMasTable)
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Warning")
            If IsNothing(p_oDTDetail) Then p_oApp.RollBackTransaction()
            Return False
        End Try

        If Not IsNothing(p_oDTDetail) Then p_oApp.CommitTransaction()

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
        loClient.Parent = "OGCallManagerHistory"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oClient = loClient
                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oOthersx.sClientNm = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")

                p_oOthersx.sAddressx = IIf(IFNull(p_oClient.Master("sHouseNox"), "") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            Else
                p_oDTMstr(0).Item("sClientID") = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sAddressx = ""
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sAddressx)
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
        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sAddressx)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getClientDetail(ByVal fnRow As Integer _
                        , ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)


        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oApp)
        loClient.Parent = "OGCallManagerHistory"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oClient = loClient
                p_oDTDetail(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oCallInfosx.xClientNmeCallInfo = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")

                p_oCallInfosx.sAddressx = IIf(IFNull(p_oClient.Master("sHouseNox"), "") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            Else
                p_oDTDetail(0).Item("sClientID") = ""
                p_oCallInfosx.xClientNmeCallInfo = ""
                p_oCallInfosx.sAddressx = ""
            End If

            RaiseEvent DetailRetrieved(fnRow, fnColDsc, p_oCallInfosx.xClientNmeCallInfo)
            RaiseEvent DetailRetrieved(fnRow, fnColDsc, p_oCallInfosx.sAddressx)
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                p_oClient = loClient
                p_oDTDetail(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oCallInfosx.xClientNmeCallInfo = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")
                p_oCallInfosx.sAddressx = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            End If
        End If

        RaiseEvent DetailRetrieved(fnRow, fnColDsc, p_oCallInfosx.xClientNmeCallInfo)
        RaiseEvent DetailRetrieved(fnRow, fnColDsc, p_oCallInfosx.sAddressx)
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub
    Private Sub InitOthers()
        p_oOthersx.sMobileNo = ""
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
    End Sub
    Private Sub InitOGCallInfo()
        p_oCallInfosx.sClientID = ""
        p_oCallInfosx.sCompnyNm = ""
        p_oCallInfosx.dCallStrt = ""
        p_oCallInfosx.dCallEndx = ""
        p_oCallInfosx.sRemarks = ""
        p_oCallInfosx.sSourceCd = ""
        p_oCallInfosx.cTLMStatx = ""
        p_oCallInfosx.xClientNmeCallInfo = ""
        p_oCallInfosx.sAddressx = ""
        p_oCallInfosx.xAgentNme = ""
    End Sub

    Private Function isEntryOk() As Boolean

        'Check validity of transaction date
        If p_oDTMstr(0).Item("dEntryDte") = "" Then
            MsgBox("Transaction date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check if application has client
        If p_oDTMstr(0).Item("sClientID") = "" Then
            MsgBox("Client Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If


        If p_oDTMstr(0).Item("sSourceNo") = "" Then
            MsgBox("Leads Info seems to have a problem! Please check your entry...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("sReferNox") = "" Then
            MsgBox("Sales Info seems to have a problem! Please check your entry...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("sRemarksx") = "" Then
            MsgBox("Note Info seems to have a problem! Please check your entry...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        p_sCallDate = ""
        p_bCallNext = False
        Call InitOthers()

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function


    Private Function getAgent(ByVal sAgentIDx As String) As String
        Dim lsSQL As String

        lsSQL = "SELECT a.sUserIDxx sUserIDxx" &
                    ", c.sCompnyNm sCompnyNm" &
                    ", a.sEmployNo sEmployNo" &
              " FROM xxxSysUser  a " &
              " LEFT JOIN Employee_Master001 b " &
              " ON a.sEmployNo = b.sEmployID " &
              " LEFT JOIN Client_Master c " &
              " ON b.sEmployID = c.sClientID " &
              " WHERE  sUserIDxx = " & strParm(sAgentIDx)

        Dim loDT As DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Return ""
        If loDT(0)("sEmployNo") <> "" Then
            Return loDT(0)("sCompnyNm")
        Else
            Return Decrypt(loDT(0)("sUserName"), "08220326")
        End If


    End Function

    Private Function getSQ_Detail() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sClientID" & _
                    ", a.sMobileNo" & _
                    ", a.sAgentIDx" & _
                    ", IFNULL (a.dCallStrt, '') dCallStrt" & _
                    ", IFNULL (a.dCallEndx, '') dCallEndx" & _
                    ", a.cTLMStatx" & _
                    ", a.sRemarksx" & _
                    ", a.sSourceCd" & _
                    ", a.sReferNox" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function
    Private Function getSQ_Master() As String
        Return "SELECT b.sCompnyNm sCompnyNm" &
                    ", a.sMobileNo sMobileNo" &
                    ", a.sClientID sClientID" &
              " FROM " & p_sMasTable & " a" &
                    ", Client_Master b" &
              " WHERE a.sClientID = b.sClientID"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sClientID sClientID" &
                    ", b.sCompnyNm sCompnyNm" &
                    ", a.sMobileNo sMobileNo" &
              " FROM " & p_sMasTable & " a" &
                    ", Client_Master b" &
              " WHERE a.sClientID = b.sClientID" &
              " GROUP BY a.sMobileNo"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "OGCallManagerHistory"
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCD = p_oApp.BranchCode
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub
    Private Class Others
        Public sMobileNo As String
        Public sClientNm As String
        Public sAddressx As String
    End Class

    Private Class OGCallInfo
        Public sTransNox As String
        Public sClientID As String
        Public sCompnyNm As String
        Public dCallStrt As String
        Public dCallEndx As String
        Public xAgentNme As String
        Public sRemarks As String
        Public sSourceCd As String
        Public cTLMStatx As String
        Public xClientNmeCallInfo As String
        Public sAddressx As String
    End Class

End Class
