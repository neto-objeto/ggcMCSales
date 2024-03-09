'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Outgoing Call Manager
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
'  Kalyptus [ 04/30/2016 09:20 pm ]
'      Started creating this object.
'  Kalyptus [ 04/30/2016 11:20 pm ]
'       Added a parent property
'
' NOTE:
'  Status of Tables subject for Call_Outgoing
'       0 - For schedule
'       1 - Scheduled 
'       2 - Done
'       3 - Cancelled
'       4 - Do Not schedule this for calling

'  Mac [ 05/02/2018 04:00 pm ]
'       Integrate Activity Inquiry
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient

Public Class OGCallManager
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_sParent As String
    Private p_bCallNext As Boolean
    Private p_sCallDate As String
    Private p_oOthersx As New Others
    Private p_oClient As ggcClient.Client
    Private p_cSubscriber As String
    Private p_cLeadSource As String = "0" '0 default, 1 lending

    Private Const p_sMasTable As String = "Call_Outgoing"
    Private Const p_sMsgHeadr As String = "Outgoing Call"
    Private Const pxeSourceCode As String = "ApCd"

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
                Select Case Index
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
                    Case "sclientnm" ' 80
                        getClient(2, 80, value, False, False)
                    Case "saddressx" ' 81
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    'leads source
    WriteOnly Property LeadSource()
        Set(value)
            p_cLeadSource = value
        End Set
    End Property

    'subscriber
    WriteOnly Property Subscriber()
        Set(value)
            p_cSubscriber = value
        End Set
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
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

    'Public Function NewTransaction()
    Public Function NewTransaction() As Boolean
        Dim lsSQL As String

        p_oDTMstr = Nothing

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

        p_sCallDate = ""
        p_bCallNext = False
        Call InitOthers()

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

        Try
            Select Case Trim(p_oDTMstr(0).Item("cTLMStatx"))
                'Case "", "AM", "WN", "NA", "UR", "NI"
                Case "PS"
                    If p_oDTMstr(0).Item("sClientID") <> "" Then
                        If p_oClient.Master("sLastName") <> "" Then
                            If Not p_oClient.SaveClient Then
                                MsgBox("Unable to save client info!", vbOKOnly, p_sMsgHeadr)
                                If p_sParent = "" Then p_oApp.RollBackTransaction()
                                Return False
                            End If

                            If p_oClient.Master("sClientID") = "" Then
                                p_oDTMstr(0).Item("sClientID") = ""
                            Else
                                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                            End If
                        End If
                    End If

                    'create entry on hotline outgoing
                    lsSQL = "INSERT INTO HotLine_Outgoing SET" & _
                                "  sTransNox = " & strParm(GetNextCode("HotLine_Outgoing", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                                ", dTransact = " & dateParm(p_oApp.SysDate) & _
                                ", sDivision = " & strParm("TLM") & _
                                ", sMobileNo = " & strParm(IIf(p_oDTMstr(0)("sMobileNo").Length = 11, "+63" & Right(p_oDTMstr(0)("sMobileNo"), 10), p_oDTMstr(0)("sMobileNo"))) & _
                                ", sMessagex = " & strParm("Guanzon Group: Your endorsement code is " & p_oDTMstr(0)("sApprovCd") & ". Present this to any Guanzon motorcycle shop on your purchase. For inquiries, contact us at " & getHotline(p_oDTMstr(0)("cSubscrbr")) & ". Thank you.") & _
                                ", cSubscrbr = " & strParm(p_oDTMstr(0)("cSubscrbr")) & _
                                ", dDueUntil = " & dateParm(DateAdd(DateInterval.Day, 5, p_oApp.SysDate)) & _
                                ", cSendStat = " & strParm(xeLogical.NO) & _
                                ", nNoRetryx = " & strParm(xeLogical.NO) & _
                                ", sUDHeader = " & strParm("") & _
                                ", sReferNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                                ", sSourceCd = " & strParm(pxeSourceCode) & _
                                ", cTranStat = " & strParm(xeLogical.NO) & _
                                ", nPriority = 1 " & _
                                ", sModified = " & strParm(p_oApp.UserID) & _
                                ", dModified = " & dateParm(p_oApp.SysDate)

                    Call p_oApp.Execute(lsSQL, "HotLine_Outgoing")

                    p_oDTMstr(0)("cSMSStatx") = "1"
                Case "NN"
                    lsSQL = "INSERT INTO HotLine_Outgoing SET" & _
                            "  sTransNox = " & strParm(GetNextCode("HotLine_Outgoing", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                            ", dTransact = " & dateParm(p_oApp.SysDate) & _
                            ", sDivision = " & strParm("TLM") & _
                            ", sMobileNo = " & strParm(IIf(p_oDTMstr(0)("sMobileNo").Length = 11, "+63" & Right(p_oDTMstr(0)("sMobileNo"), 10), p_oDTMstr(0)("sMobileNo"))) & _
                            ", sMessagex = " & strParm("Guanzon Group: Thank you for answering our call. For discounts and other concern, contact us at GLOBE-09178682713, SMART-09988577098, SUN-09258218445.") & _
                            ", cSubscrbr = " & strParm(p_oDTMstr(0)("cSubscrbr")) & _
                            ", dDueUntil = " & dateParm(DateAdd(DateInterval.Day, 5, p_oApp.SysDate)) & _
                            ", cSendStat = " & strParm(xeLogical.NO) & _
                            ", nNoRetryx = " & strParm(xeLogical.NO) & _
                            ", sUDHeader = " & strParm("") & _
                            ", sReferNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                            ", sSourceCd = " & strParm("TLMH") & _
                            ", cTranStat = " & strParm(xeLogical.NO) & _
                            ", nPriority = 1 " & _
                            ", sModified = " & strParm(p_oApp.UserID) & _
                            ", dModified = " & dateParm(p_oApp.SysDate)

                    Call p_oApp.Execute(lsSQL, "HotLine_Outgoing")

                    p_oDTMstr(0)("cSMSStatx") = "1"
                Case "UR" 'tag the client mobile as unreachable
                    lsSQL = "UPDATE Client_Mobile SET" &
                                " nUnreachx = nUnreachx + 1" &
                            " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) &
                                " AND sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

                    Call p_oApp.Execute(lsSQL, "Client_Mobile")
                Case Else
            End Select

            'tag date last called by telemarketing on client mobile
            lsSQL = "UPDATE Client_Mobile SET" &
                        " dLastCall = " & dateParm(p_oApp.SysDate) &
                    " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) &
                        " AND sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

            Call p_oApp.Execute(lsSQL, "Client_Mobile")

            'Save master table 
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)
            Else
                'UPDATE THE SOURCE TRANSACTION OF THE OUTGOING CALL
                'Determine the table
                Dim lsTableNme As String
                Select Case UCase(p_oDTMstr(0).Item("sSourceCD"))
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
                                        ", sRemarks2 = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "tlm_client"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                        ", cTranStat = '0'" &
                                        ", sAgentIDx = " & strParm(p_oApp.UserID) &
                                        ", sRemarksx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                    " WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))
                        Case "activity_inquiry"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                        ", cTranStat = '0'" &
                                        ", sAgentIDx = " & strParm(p_oApp.UserID) &
                                        ", sNotesxxx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                    " WHERE sInqryIDx = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "mc_credit_application"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET  cTLMStatx = '0'" &
                                        ", sTLMAgent = " & strParm(p_oApp.UserID) &
                                        ", dFollowUp = " & datetimeParm(p_sCallDate) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "ganado_online"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET  cTranStat = '1'" &
                                        ", sRemarksx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                        ", sTLMAgent = " & strParm(p_oApp.UserID) &
                                        ", dFollowUp = " & datetimeParm(p_sCallDate) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case Else
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET  dFollowUp = " & datetimeParm(p_sCallDate) &
                                        ", cTranStat = '0'" &
                                        ", sAgentIDx = " & strParm(p_oApp.UserID) &
                                        ", sRemarksx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                    End Select
                Else
                    'Set the cTranStat of this call as the transtat of the source transaction
                    Select Case LCase(lsTableNme)
                        Case "tlm_client"
                            lsSQL = "UPDATE " & lsTableNme &
                                        "  SET cTranStat = " & strParm(p_oDTMstr(0).Item("cTranStat")) &
                                            ", sRemarksx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "activity_inquiry"
                            lsSQL = "UPDATE " & lsTableNme & _
                                        "  SET cTranStat = " & strParm(p_oDTMstr(0).Item("cTranStat")) & _
                                            ", sNotesxxx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) & _
                                   " WHERE sInqryIDx = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "mc_product_inquiry", "mp_product_inquiry"
                            lsSQL = "UPDATE " & lsTableNme & _
                                        "  SET cTranStat = " & strParm(p_oDTMstr(0).Item("cTranStat")) & _
                                            ", sRemarks2 = " & strParm(p_oDTMstr(0).Item("sRemarksx")) & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                        Case "mc_credit_application"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET cTLMStatx = " & strParm(p_oDTMstr(0).Item("cTranStat")) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case "ganado_online"
                            lsSQL = "UPDATE " & lsTableNme &
                                    " SET cTranStat = " & strParm(p_oDTMstr(0).Item("cTranStat")) &
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sReferNox"))
                        Case Else
                            lsSQL = "UPDATE " & lsTableNme & _
                                        "  SET cTranStat = " & strParm(p_oDTMstr(0).Item("cTranStat")) & _
                                            ", sRemarksx = " & strParm(p_oDTMstr(0).Item("sRemarksx")) & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                    End Select
                End If

                Call p_oApp.Execute(lsSQL, lsTableNme)

                p_oDTMstr(0)("dModified") = p_oApp.SysDate
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID)
            End If

            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, p_sMasTable)
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "Warning")
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            Return False
        End Try

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Function GetLeadsTotal() As Integer
        Dim lsSQL As String
        Dim lsCondition As String
        Dim loDta As DataTable

        lsSQL = ""
        Select Case p_cLeadSource
            Case "0"
                lsSQL = " AND sSourceCd NOT IN ('" & pxeCOS_LENDING & _
                                                "', '" & pxeCOS_MCSALES & _
                                                "', '" & pxeCOS_CA & _
                                                "', '" & pxeCOS_MPINQR & "')"
            Case "1"
                lsSQL = " AND sSourceCd = " & strParm(pxeCOS_LENDING)
            Case "2"
                lsSQL = " AND sSourceCd IN ('" & pxeCOS_MPINQR & "', '" & pxeCOS_MCSALES & "')"
            Case "3"
                lsSQL = " AND sSourceCd = " & strParm(pxeCOS_CA)
        End Select

        Select Case p_cSubscriber
            Case "1"
                lsCondition = " AND cSubscrbr IN ('1', '2')"
            Case "3"
                lsCondition = " AND cSubscrbr IN ('1', '2', '3')"
            Case Else
                lsCondition = " AND cSubscrbr = " & strParm(p_cSubscriber)
        End Select

        lsSQL = "SELECT sTransNox, sAgentIDx" & _
                " FROM " & p_sMasTable & _
                " WHERE (cTranStat = '0'" & _
                    " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" & _
                    lsCondition & _
                    lsSQL & _
                " ORDER BY dTransact ASC, cTranStat DESC, sAgentIDx DESC"

        Debug.Print(lsSQL)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        Return loDta.Rows.Count
    End Function

    'Retrieve a scheduled call for the logged Agent
    Function Get2Call() As String
        Dim lsSQL As String
        Dim lsCondition As String
        Dim loDta As DataTable

        Select Case p_cSubscriber
            Case "1"
                lsCondition = " AND cSubscrbr IN ('1', '2')"
            Case "3"
                lsCondition = " AND cSubscrbr IN ('1', '2', '3')"
            Case Else
                lsCondition = " AND cSubscrbr = " & strParm(p_cSubscriber)
        End Select

        lsSQL = ""
        Select Case p_cLeadSource
            Case "0"
                'prioritize Ganado Online
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_GANADO) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)

                'this means that the query has a result, proceed to data processing
                If loDta.Rows.Count > 0 Then GoTo processRecord

                'prioritize MC Credit Application
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_CA) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)

                'this means that the query has a result, proceed to data processing
                If loDta.Rows.Count > 0 Then GoTo processRecord

                '2nd priority, the MC Inquiry on leads to pick
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_INQUIRY) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)

                'this means that the query has a result, proceed to data processing
                If loDta.Rows.Count > 0 Then GoTo processRecord

                'least priority other lead source
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd NOT IN ('" & pxeCOS_LENDING &
                                                "', '" & pxeCOS_MCSALES &
                                                "', '" & pxeCOS_CA &
                                                "', '" & pxeCOS_INQUIRY &
                                                "', '" & pxeCOS_MPINQR & "')" &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)
            Case "1"
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_LENDING) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"

                loDta = p_oApp.ExecuteQuery(lsSQL)
            Case "2"
                'prioritize the MP Inquiry on leads to pick
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_MPINQR) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)

                'this means that the query has a result, proceed to data processing
                If loDta.Rows.Count > 0 Then GoTo processRecord

                'second pick the leads from MC Sales
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_MCSALES) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)
            Case "3"
                lsSQL = "SELECT sTransNox, sAgentIDx" &
                        " FROM " & p_sMasTable &
                        " WHERE (cTranStat = '0'" &
                            " OR (cTranStat = '1' AND sAgentIDx = " & strParm(p_oApp.UserID) & "))" &
                            lsCondition &
                            " AND sSourceCd = " & strParm(pxeCOS_CA) &
                        " ORDER BY cSubscrbr DESC, dTransact ASC, sTransNox ASC, cTranStat DESC, sAgentIDx DESC" &
                        " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsSQL)
            Case Else
                Return ""
        End Select

processRecord:
        If loDta.Rows.Count <= 0 Then
            lsSQL = ""
        Else
            p_oApp.BeginTransaction()

            If IFNull(loDta(0).Item("sAgentIDx"), "") = "" Then
                lsSQL = "UPDATE " & p_sMasTable & _
                       " SET cTranStat = '1'" & _
                          ", sAgentIDx = " & strParm(p_oApp.UserID) & _
                       " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox"))
                If p_oApp.Execute(lsSQL, p_sMasTable) = 0 Then
                    lsSQL = ""
                Else
                    lsSQL = loDta(0).Item("sTransNox")
                End If
            Else
                lsSQL = loDta(0).Item("sTransNox")
            End If

            p_oApp.CommitTransaction()
        End If

        Return lsSQL
    End Function

    Public Function getHistory() As DataTable
        Dim loDta As DataTable
        Dim lsSQL_Ganado As String
        Dim lsSQL_MC_Inquiry As String
        Dim lsSQL_MC_Referral As String
        Dim lsSQL_Call_Incoming As String
        Dim lsSQL_SMS_Incoming As String
        Dim lsSQL_Call_Outgoing As String
        Dim lsSQL_SMS_Outgoing As String
        Dim lsSQL_MC_SO_Master As String
        Dim lsSQL_Activity_Inquiry As String
        Dim lsSQL_MC_Credit_Application As String
        Dim lsSQL As String

        If p_nEditMode <> xeEditMode.MODE_READY Then Return Nothing

        'Please make sure that sMobileNo is Index on EACH table
        lsSQL_MC_Inquiry = "SELECT a.dTransact, a.sRemarks1 sRemarksx, a.cTranStat, a.sTransNox, 'MC_Inquiry' sTableNme" & _
                          " FROM MC_Product_Inquiry a" & _
                             " LEFT JOIN Client_Mobile b ON a.sClientID = b.sClientID" & _
                          " WHERE b.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")) & _
                            " AND a.sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))

        lsSQL_MC_Referral = "SELECT a.dTransact, a.sRemarksx, a.cTranStat, a.sTransNox, 'MC_Referral' sTableNme" & _
                           " FROM MC_Referral a" & _
                              " LEFT JOIN Client_Mobile b ON a.sClientID = b.sClientID" & _
                           " WHERE b.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")) & _
                             " AND a.sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))

        lsSQL_Call_Incoming = "SELECT a.dTransact, a.sRemarksx, a.cTranStat, a.sTransNox, 'Call_Incoming' sTableNme" & _
                             " FROM Call_Incoming a WHERE a.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

        lsSQL_Call_Outgoing = "SELECT a.dTransact, a.sRemarksx, a.cTranStat, a.sTransNox, 'Call_Outgoing' sTableNme" & _
                             " FROM Call_Outgoing a WHERE a.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

        lsSQL_SMS_Incoming = "SELECT a.dTransact, a.sMessagex sRemarksx, a.cTranStat, a.sTransNox, 'SMS_Incoming' sTableNme" & _
                            " FROM SMS_Incoming a WHERE a.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

        lsSQL_SMS_Outgoing = "SELECT a.dTransact, a.sMessagex sRemarksx, a.cTranStat, a.sTransNox, 'SMS_Outgoing' sTableNme" & _
                            " FROM SMS_Outgoing a WHERE a.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo"))

        lsSQL_MC_SO_Master = "SELECT a.dTransact, a.sRemarksx sRemarksx, a.cTranStat, a.sTransNox, 'MC_SO_Master' sTableNme" & _
                            " FROM MC_SO_Master a" & _
                                " LEFT JOIN Client_Mobile b ON a.sClientID = b.sClientID" & _
                            " WHERE b.sMobileNo = " & strParm(p_oDTMstr(0).Item("sMobileNo")) & _
                             " AND a.sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))

        lsSQL_Activity_Inquiry = "SELECT" & _
                                    "  a.dInquirex dTransact" & _
                                    ", CONCAT(IFNULL(a.sColorNme, ''), ' ', IFNULL(a.sModelNme, ''), ' ', IFNULL(a.sBrandNme, '')) sRemarksx" & _
                                    ", a.cTranStat cTranStat" & _
                                    ", a.sInqryIDx sTransNox" & _
                                    ", 'Activity_Inquiry' sTableNme" & _
                            " FROM Activity_Inquiry a" & _
                                ", Call_Outgoing b" & _
                            " WHERE b.sReferNox = a.sInqryIDx" & _
                                " AND b.sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                                " AND b.sSourceCD IN ('GBF', 'FSCU', 'DC', 'OTH')"

        lsSQL_MC_Credit_Application = "SELECT" &
                                        "  a.dAppliedx dTransact" &
                                        ", 'Approved credit application' sRemarksx" &
                                        ", a.cTLMStatx cTranStat" &
                                        ", a.sTransNox sTransNox" &
                                        ", 'MC_Credit_Application' sTableNme" &
                                    " FROM MC_Credit_Application a" &
                                        ", Call_Outgoing b" &
                                    " WHERE b.sReferNox = a.sTransNox" &
                                        " AND b.sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) &
                                        " AND b.sSourceCD = 'MCCA'"

        lsSQL_Ganado = "SELECT" &
                            "  a.dCreatedx dTransact" &
                            ", 'Guanzon Ganado' sRemarksx" &
                            ", a.cTranStat" &
                            ", a.sTransNox sTransNox" &
                            ", 'Ganado_Online' sTableNme" &
                        " FROM Ganado_Online a" &
                            ", Call_Outgoing b" &
                        " WHERE b.sReferNox = a.sTransNox" &
                            " AND b.sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) &
                            " AND b.sSourceCD = " & strParm(pxeCOS_GANADO)

        lsSQL = "SELECT dTransact, sRemarksx, cTranStat, sTransNox, sTableNme" &
               " FROM (" & lsSQL_MC_Inquiry &
                 " UNION " & lsSQL_MC_Referral &
                 " UNION " & lsSQL_Call_Incoming &
                 " UNION " & lsSQL_Call_Outgoing &
                 " UNION " & lsSQL_SMS_Incoming &
                 " UNION " & lsSQL_SMS_Outgoing &
                 " UNION " & lsSQL_MC_SO_Master &
                 " UNION " & lsSQL_Activity_Inquiry &
                 " UNION " & lsSQL_Ganado &
                 " UNION " & lsSQL_MC_Credit_Application & ") x" &
               " ORDER BY dTransact DESC"

        loDta = p_oApp.ExecuteQuery(lsSQL)

        Return loDta
    End Function

    Public Function IssueCodeApproval() As String
        If p_nEditMode <> xeEditMode.MODE_READY Then Return ""

        If IFNull(p_oDTMstr(0).Item("sApprovCd")) <> "" Then Return p_oDTMstr(0).Item("sApprovCd")

        'mac 2021.05.24
        'check on other records if he was already issued approval code
        Dim lsSQL As String
        lsSQL = "SELECT sApprovCd FROM Call_Outgoing" & _
                " WHERE cTranStat <> '3" & _
                    " AND sClientID = " & strParm(p_oDTMstr(0).Item("sClientID")) & _
                    " AND sSourceCD = " & strParm(p_oDTMstr(0).Item("sSourceCD")) & _
                    " AND sReferNox = " & strParm(p_oDTMstr(0).Item("sReferNox")) & _
                " ORDER BY sTransNox DESC LIMIT 1"

        Dim loRS As DataTable = p_oApp.ExecuteQuery(lsSQL)
        If loRS.Rows.Count = 1 Then
            p_oDTMstr(0).Item("sApprovCd") = loRS(0)("sApprovCd")
        Else
            Dim oApproval As CodeApproval
            oApproval = New CodeApproval
            If p_oDTMstr(0).Item("sSourceCD") = "LEND" Then
                oApproval.XSystem = CodeApproval.pxePreApproved
            Else
                oApproval.XSystem = CodeApproval.pxeTeleMktg
            End If

            oApproval.DateRequested = CDate(p_oApp.SysDate.ToString("MM-dd-yyyy"))
            oApproval.MiscInfo = p_oOthersx.sClientNm
            oApproval.IssuedBy = "6"        '6 is the Issuee Code of Telemarketing....

            If Not oApproval.Encode() Then
                Return ""
            End If

            p_oDTMstr(0).Item("sApprovCd") = oApproval.Result
        End If
        'end - check on other records if he was already issued approval code

        p_oApp.BeginTransaction()

        'Save the approval code issued to the System_Code_Approval Table
        lsSQL = "INSERT INTO System_Code_Approval" & _
               " SET sTransNox = " & strParm(GetNextCode("System_Code_Approval", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                  ", dTransact = " & dateParm(p_oApp.SysDate) & _
                  ", sSystemCD = " & strParm(CodeApproval.pxeTeleMktg) & _
                  ", sReqstdBy = NULL" & _
                  ", dReqstdxx = " & dateParm(p_oDTMstr(0).Item("dCallStrt")) & _
                  ", cIssuedBy = " & strParm("6") & _
                  ", sMiscInfo = " & strParm(p_oOthersx.sClientNm) & _
                  ", sRemarks1 = ''" & _
                  ", sRemarks2 = ''" & _
                  ", sApprCode = " & strParm(p_oDTMstr(0).Item("sApprovCd")) & _
                  ", cTranStat = " & strParm("0") & _
                  ", sModified = " & strParm(p_oApp.UserID) & _
                  ", dModified = " & dateParm(p_oApp.SysDate)
        Call p_oApp.Execute(lsSQL, "System_Code_Approval")

        p_oApp.CommitTransaction()

        Return p_oDTMstr(0).Item("sApprovCd")
    End Function

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dcallstrt", "dcallendx"
                    p_oDTMstr(0).Item(lnCtr) = ggcAppDriver.xsNULL_DATE
                Case "nnoretryx", "nsmssentx"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case "dmodified", "smodified"
                Case "ctranstat", "csubscrbr", "ccallstat", "ctlmstatx", "csmsstatx"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
    End Sub

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
        loClient.Parent = "OGCallManager"

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
            RaiseEvent MasterRetrieved(81, p_oOthersx.sAddressx)
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

    Private Function getSQ_Master() As String
        Return "SELECT a.sTransNox" & _
                    ", a.dTransact" & _
                    ", a.sClientID" & _
                    ", a.sMobileNo" & _
                    ", a.sRemarksx" & _
                    ", a.sReferNox" & _
                    ", a.sSourceCD" & _
                    ", a.sApprovCd" & _
                    ", a.cTranStat" & _
                    ", a.sAgentIDx" & _
                    ", a.dCallStrt" & _
                    ", a.dCallEndx" & _
                    ", a.nNoRetryx" & _
                    ", a.cSubscrbr" & _
                    ", a.cCallStat" & _
                    ", a.cTLMStatx" & _
                    ", a.cSMSStatx" & _
                    ", a.nSMSSentx" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getHotline(ByVal lcSubscriber As String) As String
        Select Case lcSubscriber
            Case "0" 'globe
                Return "09178682713"
            Case "1" 'smart
                Return "09988577098"
            Case "2" 'sun
                Return "09258218445"
            Case Else 'default
                Return "09258218445"
        End Select
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "OGCallManager"
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sAddressx As String
    End Class
End Class
