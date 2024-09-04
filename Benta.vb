'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Benta Main Object
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
'  Mac [ 01/24/2024 02:24 pm ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Text

Public Class Benta
    Private Const p_sMasTable As String = "Ganado_Online"

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_sParent As String
    Private p_oClient As ggcClient.Client

    Private p_oPersonal As Personal_Info            'orig
    Private p_oPrdctInf As Product_Information      'orig
    Private p_oPaymInfo As Payment_Information      'orig
    Private p_oFinancer As Financer_Info            'orig

    Private p_xPersonal As Personal_Info            'temp
    Private p_xPrdctInf As Product_Information      'temp   
    Private p_xPaymInfo As Payment_Information      'temp
    Private p_xFinancer As Financer_Info            'temp

    Private p_oOthrInfo As Other_Info

    Public Event PersonalRetreived(ByVal Index As Integer, ByVal Value As Object)
    Public Event ProductRetreived(ByVal Index As Integer, ByVal Value As Object)
    Public Event FinancerRetreived(ByVal Index As Integer, ByVal Value As Object)
    Public Event PaymentRetreived(ByVal Index As Integer, ByVal Value As Object)

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

    Public Property Personal_Info(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                With p_xPersonal
                    Select Case LCase(Index)
                        Case "slastname"
                            Return .sLastName
                        Case "sfrstname"
                            Return .sFrstName
                        Case "smiddname"
                            Return .sMiddName
                        Case "ssuffixnm"
                            Return .sSuffixNm
                        Case "smaidennm"
                            Return .sMaidenNm
                        Case "cgendercd"
                            Return .cGenderCd
                        Case "dbirthdte"
                            Return .dBirthDte
                        Case "sbirthplc"
                            Return getTown(.sBirthPlc,, True)
                        Case "shousenox"
                            Return .sHouseNox
                        Case "saddressx"
                            Return .sAddressx
                        Case "stownidxx"
                            Return getTown(.sTownIDxx,, True)
                        Case "smobileno"
                            Return .sMobileNo
                        Case "semailadd"
                            Return .sEmailAdd
                        Case Else
                            Return vbEmpty
                    End Select
                End With
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            With p_xPersonal
                Select Case LCase(Index)
                    Case "slastname"
                        .sLastName = value
                        RaiseEvent PersonalRetreived(1, .sLastName)
                    Case "sfrstname"
                        .sFrstName = value
                        RaiseEvent PersonalRetreived(2, .sFrstName)
                    Case "smiddname"
                        .sMiddName = value
                        RaiseEvent PersonalRetreived(3, .sMiddName)
                    Case "ssuffixnm"
                        .sSuffixNm = value
                        RaiseEvent PersonalRetreived(4, .sSuffixNm)
                    Case "smaidennm"
                        .sMaidenNm = value
                        RaiseEvent PersonalRetreived(5, .sMaidenNm)
                    Case "cgendercd"
                        .cGenderCd = value
                        RaiseEvent PersonalRetreived(6, .cGenderCd)
                    Case "dbirthdte"
                        .dBirthDte = value
                        RaiseEvent PersonalRetreived(7, .dBirthDte)
                    Case "sbirthplc"
                        .sBirthPlc = value
                        RaiseEvent PersonalRetreived(8, getTown(.sBirthPlc, True, False))
                    Case "shousenox"
                        .sHouseNox = value
                        RaiseEvent PersonalRetreived(9, .sHouseNox)
                    Case "saddressx"
                        .sAddressx = value
                        RaiseEvent PersonalRetreived(10, .sAddressx)
                    Case "stownidxx"
                        .sTownIDxx = value
                        RaiseEvent PersonalRetreived(11, getTown(.sTownIDxx, True, False))
                    Case "smobileno"
                        .sMobileNo = value
                        RaiseEvent PersonalRetreived(12, .sMobileNo)
                    Case "semailadd"
                        .sEmailAdd = value
                        RaiseEvent PersonalRetreived(13, .sEmailAdd)
                End Select
            End With
        End Set
    End Property

    Public Property Product_Info(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                With p_xPrdctInf
                    Select Case LCase(Index)
                        Case "sbrandidx"
                            Return getMCBrand(.sBrandIDx,, True)
                        Case "smodelidx"
                            Return getMCModel(.sModelIDx,, True)
                        Case "scoloridx"
                            Return getColor(.sColorIDx,, True)
                        Case "nselprice"
                            Return .nSelPrice
                        Case "dpricexxx"
                            Return .dPricexxx
                        Case Else
                            Return vbEmpty
                    End Select
                End With
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            With p_xPrdctInf
                Select Case LCase(Index)
                    Case "sbrandidx"
                        .sBrandIDx = value
                        RaiseEvent ProductRetreived(0, getMCBrand(.sBrandIDx, True, False))
                    Case "smodelidx"
                        .sModelIDx = value
                        RaiseEvent ProductRetreived(1, getMCModel(.sModelIDx, True, False))
                    Case "scoloridx"
                        .sColorIDx = value
                        RaiseEvent ProductRetreived(2, getColor(.sColorIDx, True, False))
                End Select
            End With
        End Set
    End Property

    Public Property Payment_Info(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                With p_xPaymInfo
                    Select Case LCase(Index)
                        Case "stermidxx"
                            Return .sTermIDxx
                        Case "ndownpaym"
                            Return .nDownPaym
                        Case Else
                            Return vbEmpty
                    End Select
                End With
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            With p_xPaymInfo
                Select Case LCase(Index)
                    Case "stermidxx"
                        .sTermIDxx = value
                        RaiseEvent PaymentRetreived(2, .sTermIDxx)
                    Case "ndownpaym"
                        .nDownPaym = value
                        RaiseEvent PaymentRetreived(3, .nDownPaym)
                End Select
            End With
        End Set
    End Property

    Public Property Financer_Info(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                With p_xFinancer
                    Select Case LCase(Index)
                        Case "slastname"
                            Return .sLastName
                        Case "sfrstname"
                            Return .sFrstName
                        Case "smiddname"
                            Return .sMiddName
                        Case "ssuffixnm"
                            Return .sSuffixNm
                        Case "saddressx"
                            Return .sAddressx
                        Case "scntrycde"
                            Return .sCntryCde
                        Case "smobileno"
                            Return .sMobileNo
                        Case "swhatsapp"
                            Return .sWhatsApp
                        Case "swchatapp"
                            Return .sWChatApp
                        Case "sfbaccntx"
                            Return .sFbAccntx
                        Case "semailadd"
                            Return .sEmailAdd
                        Case "sfincomex"
                            Return .sFIncomex
                        Case "sreltionx"
                            Return .sReltionx
                        Case Else
                            Return vbEmpty
                    End Select
                End With
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            With p_xFinancer
                Select Case LCase(Index)
                    Case "slastname"
                        .sLastName = value
                        RaiseEvent FinancerRetreived(1, .sLastName)
                    Case "sfrstname"
                        .sFrstName = value
                        RaiseEvent FinancerRetreived(2, .sFrstName)
                    Case "smiddname"
                        .sMiddName = value
                        RaiseEvent FinancerRetreived(3, .sMiddName)
                    Case "ssuffixnm"
                        .sSuffixNm = value
                        RaiseEvent FinancerRetreived(4, .sSuffixNm)
                    Case "saddressx"
                        .sAddressx = value
                        RaiseEvent FinancerRetreived(5, .sAddressx)
                    Case "scntrycde"
                        .sCntryCde = value
                        RaiseEvent FinancerRetreived(6, .sCntryCde)
                    Case "smobileno"
                        .sMobileNo = value
                        RaiseEvent FinancerRetreived(7, .sMobileNo)
                    Case "swhatsapp"
                        .sWhatsApp = value
                        RaiseEvent FinancerRetreived(8, .sWhatsApp)
                    Case "swchatapp"
                        .sWChatApp = value
                        RaiseEvent FinancerRetreived(9, .sWChatApp)
                    Case "sfbaccntx"
                        .sFbAccntx = value
                        RaiseEvent FinancerRetreived(10, .sFbAccntx)
                    Case "semailadd"
                        .sEmailAdd = value
                        RaiseEvent FinancerRetreived(11, .sEmailAdd)
                    Case "sfincomex"
                        .sFIncomex = value
                        RaiseEvent FinancerRetreived(12, .sFIncomex)
                    Case "sreltionx"
                        .sReltionx = value
                        RaiseEvent FinancerRetreived(13, .sReltionx)
                End Select
            End With
        End Set
    End Property



    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "stransnox", "dtransact", "sclientnm", "dtargetxx", "dfollowup", "sremarksx", "sreferdby", "dcreatedx", "nlatitude", "nlongitud", "sclientid", "stlmagent", "smobileno", "semailadd", "xrefernme"
                        Return p_oDTMstr(0).Item(Index)
                    Case "cganadotp"
                        Return p_oDTMstr(0).Item("xGanadoTp")
                    Case "ctranstat"
                        Return p_oDTMstr(0).Item("xTranStat")
                    Case "csourcexx"
                        Return p_oDTMstr(0).Item("xSourcexx")
                    Case "cpaymform"
                        Return p_oDTMstr(0).Item("cPaymForm")
                    Case "srelatnid"
                    Case "scltinfox"
                        Return vbEmpty
                    Case "sfinancex"
                        Return vbEmpty
                    Case "sprdctinf"
                        Return vbEmpty
                    Case "spayminfo"
                        Return vbEmpty
                    Case Else
                        Return vbEmpty
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "cpaymform"
                        p_oDTMstr(0).Item("cPaymForm") = value
                        RaiseEvent PaymentRetreived(1, p_oDTMstr(0).Item("cPaymForm"))
                End Select
            End If
        End Set
    End Property

    Public Function PickTransaction(ByVal fsValue As String _
                                        , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String = getSQ_Master()

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sClientNm") Then Return True
            End If
        End If

        lsSQL = AddCondition(lsSQL, "a.cTranStat = '0' AND ISNULL(a.sTLMAgent)")

        If LCase(p_oApp.ProductID) = "telemktg" Then
            lsSQL = AddCondition(getSQ_Master, "a.cPaymForm = '0' AND a.cTranStat <> '3'")
        ElseIf LCase(p_oApp.ProductID) = "lrtrackr" Then
            lsSQL = AddCondition(getSQ_Master, "a.cPaymForm = '1' AND a.cTranStat <> '3'")
        Else
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            MsgBox("Application is not allowed to load benta transactions.", vbInformation, "Notice")
            Return False
        End If

        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "a.sClientNm LIKE " & strParm(fsValue & "%")
        End If

        lsSQL += " LIMIT 1"

        Dim loDT As DataTable
        loDT = New DataTable
        Debug.Print(lsSQL)
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            MsgBox("No benta transaction to load at this time.", vbInformation, "Notice")
            Return False
        Else
            Return OpenTransaction(loDT(0).Item("sTransNox"))
        End If
    End Function

    Private Function isEntryOK() As Boolean
        If p_xPersonal.sLastName = "" Then
            MsgBox("Personal information last name must not be empty.")
            Return False
        End If

        If p_xPersonal.sFrstName = "" Then
            MsgBox("Personal information first name must not be empty.")
            Return False
        End If

        If p_xPersonal.sMaidenNm = "" Then
            MsgBox("Personal information mother's maiden name must not be empty.")
            Return False
        End If

        If p_xPersonal.dBirthDte = "" Then
            MsgBox("Personal information birthday must not be null.")
            Return False
        End If

        If p_xPersonal.sBirthPlc = "" Then
            MsgBox("Personal information birth place must not be empty.")
            Return False
        End If

        If p_xPersonal.sAddressx = "" Then
            MsgBox("Personal address must not be empty.")
            Return False
        End If

        If p_xPersonal.sTownIDxx = "" Then
            MsgBox("Personal town/city must not be empty.")
            Return False
        End If

        If p_xPrdctInf.sBrandIDx = "" Then
            MsgBox("Product brand must not be empty.")
            Return False
        End If

        If p_xPrdctInf.sModelIDx = "" Then
            MsgBox("Product model must not be empty.")
            Return False
        End If

        If p_xPrdctInf.sColorIDx = "" Then
            MsgBox("Product color must not be empty.")
            Return False
        End If

        Return True
    End Function

    Public Function SaveTransaction(Optional ByVal fcTranStat As String = "") As Boolean
        If p_nEditMode <> xeEditMode.MODE_UPDATE Then
            MsgBox("Transaction is not in update mode.", vbInformation, "Notice")
            Return False
        End If

        Dim lsSQL As String = ""
        Dim lsClientID As String = ""
        Dim lsClientNm As String = ""

        Dim loCloud As GCloud

        loCloud = New GCloud

        loCloud.UserID = p_oApp.UserID

        p_oApp.BeginTransaction()

        'verified must be validated
        If (fcTranStat = "1") Then
            If Not isEntryOK() Then Return False

            lsClientID = GetNextCode("Client_Master", "sClientID", True, p_oApp.Connection, True, p_oApp.BranchCode)

            lsSQL = "INSERT INTO Client_Master SET " +
                                    "  sClientID = " + strParm(lsClientID) +
                                    ", sLastName = " + strParm(p_xPersonal.sLastName) +
                                    ", sFrstName = " + strParm(p_xPersonal.sFrstName) +
                                    ", sMiddName = " + strParm(p_xPersonal.sMiddName) +
                                    ", sMaidenNm = " + strParm(p_xPersonal.sMiddName) +
                                    ", sSuffixNm = " + strParm(p_xPersonal.sSuffixNm) +
                                    ", cGenderCd = " + strParm(p_xPersonal.cGenderCd) +
                                    ", cCvilStat = '0'" +
                                    ", sCitizenx = '01'" +
                                    ", dBirthDte = " + dateParm(p_xPersonal.dBirthDte) +
                                    ", sBirthPlc = " + strParm(p_xPersonal.sBirthPlc) +
                                    ", sHouseNox = " + strParm(p_xPersonal.sHouseNox) +
                                    ", sAddressx = " + strParm(p_xPersonal.sAddressx) +
                                    ", sTownIDxx = " + strParm(p_xPersonal.sTownIDxx) +
                                    ", sBrgyIDxx = ''" +
                                    ", sPhoneNox = ''" +
                                    ", sMobileNo = " + strParm(p_xPersonal.sMobileNo) +
                                    ", sEmailAdd = " + strParm(p_xPersonal.sEmailAdd) +
                                    ", cEducLevl = '6'" +
                                    ", sRelgnIDx = ''" +
                                    ", sTaxIDNox = ''" +
                                    ", sSSSNoxxx = ''" +
                                    ", sAddlInfo = ''" +
                                    ", sCompnyNm = " + strParm(lsClientNm) +
                                    ", sOccptnID = ''" +
                                    ", sOccptnOT = ''" +
                                    ", nGrssIncm = 0" +
                                    ", sClientNo = ''" +
                                    ", sSpouseID = ''" +
                                    ", sFatherID = ''" +
                                    ", sMotherID = ''" +
                                    ", sSiblngID = ''" +
                                    ", cClientTp = '0'" +
                                    ", cLRClient = '0'" +
                                    ", cMCClient = '0'" +
                                    ", cSCClient = '0'" +
                                    ", cSPClient = '0'" +
                                    ", cCPClient = '0'" +
                                    ", cRecdStat = '1'" +
                                    ", sModified = " + strParm(p_oApp.UserID) +
                                    ", dModified = " + datetimeParm(p_oApp.SysDate)

            If p_oApp.ExecuteActionQuery(lsSQL) <= 0 Then
                MsgBox("Unable to save client info.", vbCritical, "Warning")
                p_oApp.RollBackTransaction()
                Return False
            End If
            loCloud.AddStatement(lsSQL, SQLCondition.xeEquals, 1, True, "Client_Master", p_oApp.BranchCode, "")

            lsSQL = "INSERT INTO Client_Mobile SET" +
                    "  sClientID = " + strParm(lsClientID) +
                    ", nEntryNox = '1'" +
                    ", sMobileNo = " + strParm(p_xPersonal.sMobileNo) +
                    ", nPriority = 1" +
                    ", cSubscrbr = NULL" +
                    ", cRecdStat = '1'"

            If p_oApp.ExecuteActionQuery(lsSQL) <= 0 Then
                MsgBox("Unable to save client info.", vbCritical, "Warning")
                p_oApp.RollBackTransaction()
                Return False
            End If
            loCloud.AddStatement(lsSQL, SQLCondition.xeEquals, 1, True, "Client_Mobile", p_oApp.BranchCode, "")
        End If

        Dim lsPersonal As String = ""
        Dim lsProductx As String = ""
        Dim lsFinancer As String = ""
        Dim lsPaymentx As String = ""

        lsClientNm = p_oDTMstr(0).Item("sClientNm")

        'compare personal info
        If p_xPersonal.sLastName <> p_oPersonal.sLastName Or
            p_xPersonal.sFrstName <> p_oPersonal.sFrstName Or
            p_xPersonal.sMiddName <> p_oPersonal.sMiddName Or
            p_xPersonal.sSuffixNm <> p_oPersonal.sSuffixNm Or
            p_xPersonal.sMaidenNm <> p_oPersonal.sMaidenNm Or
            p_xPersonal.cGenderCd <> p_oPersonal.cGenderCd Or
            p_xPersonal.dBirthDte <> p_oPersonal.dBirthDte Or
            p_xPersonal.sBirthPlc <> p_oPersonal.sBirthPlc Or
            p_xPersonal.sHouseNox <> p_oPersonal.sHouseNox Or
            p_xPersonal.sAddressx <> p_oPersonal.sAddressx Or
            p_xPersonal.sTownIDxx <> p_oPersonal.sTownIDxx Or
            p_xPersonal.sMobileNo <> p_oPersonal.sMobileNo Or
            p_xPersonal.sEmailAdd <> p_oPersonal.sEmailAdd Then

            lsPersonal = json_encode(p_xPersonal)

            lsClientNm = p_xPersonal.sLastName & ", " & p_xPersonal.sFrstName

            If p_xPersonal.sSuffixNm <> "" Then
                lsClientNm += " " & p_xPersonal.sSuffixNm
            End If

            lsClientNm += " " & p_xPersonal.sMiddName
            lsClientNm = Trim(lsClientNm)
        End If

        'compare product info
        If p_xPrdctInf.sBrandIDx <> p_oPrdctInf.sBrandIDx Or
                p_xPrdctInf.sModelIDx <> p_oPrdctInf.sModelIDx Or
                p_xPrdctInf.sColorIDx <> p_oPrdctInf.sColorIDx Or
                p_xPrdctInf.nSelPrice <> p_oPrdctInf.nSelPrice Or
                p_xPrdctInf.dPricexxx <> p_oPrdctInf.dPricexxx Then
            lsProductx = json_encode(p_xPrdctInf)
        End If

        'compare financer info
        If p_xFinancer.sLastName <> p_oFinancer.sLastName Or
            p_xFinancer.sFrstName <> p_oFinancer.sFrstName Or
            p_xFinancer.sMiddName <> p_oFinancer.sMiddName Or
            p_xFinancer.sSuffixNm <> p_oFinancer.sSuffixNm Or
            p_xFinancer.sAddressx <> p_oFinancer.sAddressx Or
            p_xFinancer.sCntryCde <> p_oFinancer.sCntryCde Or
            p_xFinancer.sMobileNo <> p_oFinancer.sMobileNo Or
            p_xFinancer.sWhatsApp <> p_oFinancer.sWhatsApp Or
            p_xFinancer.sWChatApp <> p_oFinancer.sWChatApp Or
            p_xFinancer.sFbAccntx <> p_oFinancer.sFbAccntx Or
            p_xFinancer.sEmailAdd <> p_oFinancer.sEmailAdd Or
            p_xFinancer.sFIncomex <> p_oFinancer.sFIncomex Or
            p_xFinancer.sReltionx <> p_oFinancer.sReltionx Then
            lsFinancer = json_encode(p_xFinancer)
        End If

        'compare payment info
        If p_oDTMstr(0).Item("cPaymForm") = "0" Then
            p_xPaymInfo.sTermIDxx = ""
            p_xPaymInfo.nDownPaym = 0
        End If

        If p_xPaymInfo.sTermIDxx <> p_oPaymInfo.sTermIDxx Or
            p_xPaymInfo.nDownPaym <> p_oPaymInfo.nDownPaym Then
            lsPaymentx = json_encode(p_xPaymInfo)
        End If

        lsSQL = "UPDATE " & p_sMasTable & " SET " &
                    "  cPaymForm = " & strParm(p_oDTMstr(0).Item("cPaymForm")) &
                    ", sClientID = " & strParm(lsClientID)

        If lsPersonal <> "" Then
            lsSQL += ", sCltInfoF = '" & lsPersonal & "'"
        End If

        If lsFinancer <> "" Then
            lsSQL += ", sFinanceF = '" & lsFinancer & "'"
        End If

        If lsProductx <> "" Then
            lsSQL += ", sPrdctxxF = '" & lsProductx & "'"
        End If

        If lsPaymentx <> "" Then
            lsSQL += ", sPaymInfF = '" & lsPaymentx & "'"
        End If

        If (fcTranStat = "1") Then
            lsSQL += ", sClientNm = " & strParm(lsClientNm)
        End If

        lsSQL += ", cTranStat = " & strParm(fcTranStat) &
                    ", sModified = " & strParm(p_oApp.UserID) &
                    ", dModified = " & datetimeParm(p_oApp.SysDate)

        lsSQL = AddCondition(lsSQL, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

        If p_oApp.ExecuteActionQuery(lsSQL) <= 0 Then
            MsgBox("Unable to update transaction.", vbCritical, "Warning")
            p_oApp.RollBackTransaction()
            Return False
        End If
        loCloud.AddStatement(lsSQL, SQLCondition.xeEquals, 1, False)

        loCloud.CommitTransaction()

        If (loCloud.Execute) Then
            p_oApp.CommitTransaction()
            Return True
        Else
            MsgBox("Unable to sumbit transaction update.", vbCritical, "Warning")
            p_oApp.RollBackTransaction()
            Return False
        End If
    End Function

    Public Function SearchTransaction(ByVal fsValue As String _
                                        , Optional ByVal fbByCode As Boolean = False _
                                        , Optional ByVal fsUserID As String = "") As Boolean

        Dim lsSQL As String = getSQ_Master()

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sClientNm") Then Return True
            End If
        End If

        If fsUserID <> "" Then
            lsSQL = AddCondition(lsSQL, "a.sTLMAgent = " + strParm(fsUserID))
        End If

        Dim lsFilter As String
            If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "a.sClientNm LIKE " & strParm(fsValue & "%")
        End If

        Debug.Print(lsSQL)
        Dim loDT As DataTable = p_oApp.ExecuteQuery(AddCondition(lsSQL, lsFilter))

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sClientNm»dCreatedx»xTranStat" _
                                        , "Client»Date»Status",
                                        , "a.sClientNm»a.dCreatedx»xTranStat" _
                                        , IIf(fbByCode, 1, 2))

        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function


    Private Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))

        Debug.Print(lsSQL)
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        If p_oDTMstr(0).Item("sCltInfox") <> "" Then json_decode(New Personal_Info, p_oDTMstr(0).Item("sCltInfox"))
        If p_oDTMstr(0).Item("sPrdctInf") <> "" Then json_decode(New Product_Information, p_oDTMstr(0).Item("sPrdctInf"))
        If p_oDTMstr(0).Item("sPaymInfo") <> "" Then json_decode(New Payment_Information, p_oDTMstr(0).Item("sPaymInfo"))
        If p_oDTMstr(0).Item("sFinancex") <> "" Then json_decode(New Financer_Info, p_oDTMstr(0).Item("sFinancex"))

        If p_oDTMstr(0).Item("cTranStat") = "0" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            If p_oDTMstr(0).Item("sTLMAgent") = "" Then
                lsSQL = "UPDATE " & p_sMasTable + " SET sTLMAgent = " + strParm(p_oApp.UserID) + " WHERE sTransNox = " + strParm(fsTransNox)

                If (p_oApp.Execute(lsSQL, p_sMasTable) <= 0) Then
                    MsgBox("Unable to update MASTER TABLE.", vbCritical, "Warning")
                    p_nEditMode = xeEditMode.MODE_UNKNOWN
                    Return False
                End If
            End If

            p_nEditMode = xeEditMode.MODE_UPDATE
        Else
            p_nEditMode = xeEditMode.MODE_READY
        End If

        Return True
    End Function

    Private Function getSQ_Master() As String
        Return "SELECT" +
                    "  a.sTransNox" +
                    ", a.dTransact" +
                    ", a.sClientNm" +
                    ", a.cSourcexx" +
                    ", a.cGanadoTp" +
                    ", a.cPaymForm" +
                    ", IF(IFNULL(a.sCltInfoF, '') = '', IFNULL(a.sCltInfox, ''), a.sCltInfoF) sCltInfox" +
                    ", IF(IFNULL(a.sFinanceF, '') = '', IFNULL(a.sFinancex, ''), a.sFinanceF) sFinancex" +
                    ", IF(IFNULL(a.sPrdctxxF, '') = '', IFNULL(a.sPrdctInf, ''), a.sPrdctxxF) sPrdctInf" +
                    ", IF(IFNULL(a.sPaymInfF, '') = '', IFNULL(a.sPaymInfo, ''), a.sPaymInfF) sPaymInfo" +
                    ", IFNULL(a.dTargetxx, '1900-01-01') dTargetxx" +
                    ", a.dFollowUp" +
                    ", a.sRemarksx" +
                    ", a.sReferdBy" +
                    ", a.sRelatnID" +
                    ", a.dCreatedx" +
                    ", a.nLatitude" +
                    ", a.nLongitud" +
                    ", IFNULL(a.sClientID, '') sClientID" +
                    ", IFNULL(a.sTLMAgent, '') sTLMAgent" +
                    ", a.cTranStat" +
                    ", CASE a.cTranStat" +
                        " WHEN '0' THEN 'Pending'" +
                        " WHEN '1' THEN 'Verified'" +
                        " WHEN '2' THEN 'Unable to Verify'" +
                        " WHEN '3' THEN 'Expired'" +
                        " WHEN '4' THEN 'Bought'" +
                        " WHEN '5' THEN 'Pending Incentive Release'" +
                        " WHEN '6' THEN 'Incentive Released'" +
                        " ELSE '-'" +
                    " END xTranStat" +
                    ", CASE a.cSourcexx" +
                        " WHEN '0' THEN 'Guanzon Circle'" +
                        " WHEN '1' THEN 'Guanzon Connect'" +
                        " WHEN '2' THEN 'Guanzon Sales Kit'" +
                        " ELSE '-'" +
                    " END xSourcexx" +
                    ", CASE a.cGanadoTp" +
                        " WHEN '1' THEN 'Motorcycle'" +
                        " WHEN '2' THEN 'Automobile'" +
                        " ELSE '-'" +
                    " END xGanadoTp" +
                    ", b.sMobileNo" +
                    ", b.sEmailAdd" +
                    ", IFNULL(c.sCompnyNm, b.sUserName) xReferNme" +
                " FROM " + p_sMasTable + " a" +
                    " LEFT JOIN App_User_Master b ON a.sReferdBy = b.sUserIDxx" +
                    " LEFT JOIN Client_Master c ON b.sEmployNo = c.sClientID" +
                " ORDER BY a.dCreatedx"
    End Function

    Private Function getSQ_MCBrand() As String
        Return "SELECT" +
                    "  sBrandIDx" +
                    ", sBrandNme" +
                " FROM Brand" +
                " WHERE cRecdStat = '1'"
    End Function

    Private Function getSQ_MCModel() As String
        Return "SELECT" +
                    "  a.sModelIDx" +
                    ", a.sModelCde" +
                    ", a.sModelNme" +
                    ", a.sBrandIDx" +
                    ", IFNULL(b.sBrandNme, '') sBrandNme" +
                    ", IFNULL(d.nSelPrice, 0.00) nSelPrice" +
                    ", IFNULL(d.dPricexxx, '1900-01-01') dPricexxx" +
                " FROM MC_Model a" +
                    " LEFT JOIN Brand b ON a.sBrandIDx = b.sBrandIDx" +
                    " LEFT JOIN MC_Model_Price d ON a.sModelIDx = d.sModelIDx" +
                    ", MC_Inventory c" +
                " WHERE a.sModelIDx = c.sModelIDx" +
                    " AND c.sBranchCd IN ('M0W1', 'M029')" +
                    " And a.cRecdStat = '1'" +
                    " AND a.cEndOfLfe = '0'" +
                " GROUP BY a.sModelIDx"
    End Function

    Private Function getSQ_Color() As String
        Return "SELECT" +
                    "  a.sColorIDx" +
                    ", a.sColorNme" +
                " FROM Color a" +
                    ", MC_Inventory b" +
                " WHERE a.sColorIDx = b.sColorIDx" +
                    " AND b.sModelIDx = " + strParm(p_xPrdctInf.sModelIDx) +
                    " AND b.sBranchCd IN ('M0W1', 'M029')" +
                " GROUP BY a.sColorIDx"
    End Function

    Private Function getSQ_TownCity() As String
        Return "SELECT" +
                    "  a.sTownIDxx" +
                    ", a.sTownName" +
                    ", b.sProvName" +
                " FROM TownCity a" +
                    " LEFT JOIN Province b ON a.sProvIDxx = b.sProvIDxx"
    End Function



    Private Sub json_decode(ByVal foJSONObject As Object,
                            ByVal fsJSONValue As String)

        Dim loSettings As New JsonSerializerSettings
        loSettings.DefaultValueHandling = DefaultValueHandling.Populate

        If (TypeOf foJSONObject Is Personal_Info) Then
            p_oPersonal = JsonConvert.DeserializeObject(Of Personal_Info)(fsJSONValue, loSettings)
            p_xPersonal = JsonConvert.DeserializeObject(Of Personal_Info)(fsJSONValue, loSettings)
        ElseIf (TypeOf foJSONObject Is Product_Information) Then
            p_oPrdctInf = JsonConvert.DeserializeObject(Of Product_Information)(fsJSONValue, loSettings)
            p_xPrdctInf = JsonConvert.DeserializeObject(Of Product_Information)(fsJSONValue, loSettings)
        ElseIf (TypeOf foJSONObject Is Payment_Information) Then
            p_oPaymInfo = JsonConvert.DeserializeObject(Of Payment_Information)(fsJSONValue, loSettings)
            p_xPaymInfo = JsonConvert.DeserializeObject(Of Payment_Information)(fsJSONValue, loSettings)
        ElseIf (TypeOf foJSONObject Is Financer_Info) Then
            p_oFinancer = JsonConvert.DeserializeObject(Of Financer_Info)(fsJSONValue, loSettings)
            p_xFinancer = JsonConvert.DeserializeObject(Of Financer_Info)(fsJSONValue, loSettings)
        End If
    End Sub

    Private Function json_encode(ByVal foJSONObject As Object) As String
        Dim loSettings As New JsonSerializerSettings

        loSettings.NullValueHandling = NullValueHandling.Ignore
        loSettings.DefaultValueHandling = DefaultValueHandling.Ignore

        Return JsonConvert.SerializeObject(foJSONObject, loSettings)
    End Function

    Private Function getMCBrand(ByVal sValue As String,
                                Optional ByVal bSearch As Boolean = False,
                                Optional ByVal bByCode As Boolean = False) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getMCBrand"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sBrandNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sBrandNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sBrandIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQ_MCBrand, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            p_xPrdctInf.sBrandIDx = loDT(0)("sBrandIDx")
            getMCBrand = loDT(0)("sBrandNme")
        Else
            loDataRow = KwikSearch(p_oApp,
                                lsSQL,
                                "",
                                "sBrandIDx»sBrandNme",
                                "ID»Brand",
                                "",
                                "sBrandIDx»sBrandNme",
                                1)

            If Not IsNothing(loDataRow) Then
                p_xPrdctInf.sBrandIDx = loDataRow("sBrandIDx")
                getMCBrand = loDataRow("sBrandNme")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        p_xPrdctInf.sBrandIDx = ""
        getMCBrand = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getMCModel(ByVal sValue As String,
                                Optional ByVal bSearch As Boolean = False,
                                Optional ByVal bByCode As Boolean = False) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getMCModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQ_MCModel, lsCondition)

        'filter selected brand
        If (p_xPrdctInf.sBrandIDx <> "") Then
            lsSQL = AddCondition(lsSQL, "a.sBrandIDx = " + strParm(p_xPrdctInf.sBrandIDx))
        End If

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            p_xPrdctInf.sBrandIDx = loDT(0)("sBrandIDx")
            RaiseEvent ProductRetreived(0, loDT(0)("sBrandNme"))
            RaiseEvent ProductRetreived(3, loDT(0)("sModelCde"))

            p_xPrdctInf.sModelIDx = loDT(0)("sModelIDx")
            p_xPrdctInf.nSelPrice = loDT(0)("nSelPrice")
            p_xPrdctInf.dPricexxx = loDT(0)("dPricexxx")
            RaiseEvent ProductRetreived(2, getColor(p_xPrdctInf.sColorIDx,, True))
            RaiseEvent ProductRetreived(7, p_xPrdctInf.nSelPrice)

            getMCModel = loDT(0)("sModelNme")
        Else
            loDataRow = KwikSearch(p_oApp,
                                lsSQL,
                                "",
                                "sModelIDx»sModelNme»sModelCde",
                                "ID»Model»Code",
                                "",
                                "sModelIDx»sModelNme»sModelCde",
                                1)

            If Not IsNothing(loDataRow) Then
                p_xPrdctInf.sBrandIDx = loDataRow("sBrandIDx")
                RaiseEvent ProductRetreived(0, loDataRow("sBrandNme"))
                RaiseEvent ProductRetreived(3, loDataRow("sModelCde"))

                p_xPrdctInf.sModelIDx = loDataRow("sModelIDx")
                p_xPrdctInf.nSelPrice = loDataRow("nSelPrice")
                p_xPrdctInf.dPricexxx = loDataRow("dPricexxx")
                RaiseEvent ProductRetreived(2, getColor(p_xPrdctInf.sColorIDx,, True))
                RaiseEvent ProductRetreived(7, p_xPrdctInf.nSelPrice)

                getMCModel = loDT(0)("sModelNme")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        p_xPrdctInf.sColorIDx = ""
        p_xPrdctInf.sModelIDx = ""
        p_xPrdctInf.nSelPrice = 0
        p_xPrdctInf.dPricexxx = ""

        RaiseEvent ProductRetreived(3, "")
        RaiseEvent ProductRetreived(2, "")
        RaiseEvent ProductRetreived(7, "")
        getMCModel = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getColor(ByVal sValue As String,
                                Optional ByVal bSearch As Boolean = False,
                                Optional ByVal bByCode As Boolean = False) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getColor"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sColorNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sColorNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sColorIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQ_Color, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            p_xPrdctInf.sColorIDx = loDT(0)("sColorIDx")
            getColor = loDT(0)("sColorNme")
        Else
            loDataRow = KwikSearch(p_oApp,
                                lsSQL,
                                "",
                                "sColorIDx»sColorNme",
                                "ID»Color",
                                "",
                                "a.sColorIDx»a.sColorNme",
                                1)

            If Not IsNothing(loDataRow) Then
                p_xPrdctInf.sColorIDx = loDataRow("sColorIDx")
                getColor = loDataRow("sColorNme")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        p_xPrdctInf.sColorIDx = ""
        getColor = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getTown(ByRef sValue As String,
                                Optional ByVal bSearch As Boolean = False,
                                Optional ByVal bByCode As Boolean = False) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getTown"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sTownName LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sTownName = " & strParm(sValue)
                End If
            Else
                lsCondition = "sTownIDxx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQ_TownCity, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            sValue = loDT(0)("sTownIDxx")
            getTown = loDT(0)("sTownName") + ", " + loDT(0)("sProvName")
        Else
            loDataRow = KwikSearch(p_oApp,
                                lsSQL,
                                "",
                                "sTownIDxx»sTownName»sProvName",
                                "ID»Town»Province",
                                "",
                                "sTownIDxx»sTownName»sProvName",
                                1)

            If Not IsNothing(loDataRow) Then
                sValue = loDataRow("sTownIDxx")
                getTown = loDataRow("sTownName") + ", " + loDataRow("sProvName")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        sValue = ""
        getTown = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "Benta"

        p_oOthrInfo = New Other_Info

        p_xPersonal = New Personal_Info
        p_xPrdctInf = New Product_Information
        p_xPaymInfo = New Payment_Information
        p_xFinancer = New Financer_Info

        p_oPersonal = New Personal_Info
        p_oPrdctInf = New Product_Information
        p_oPaymInfo = New Payment_Information
        p_oFinancer = New Financer_Info

        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Class Other_Info
        Property sBrandNme As String
        Property sModelNme As String
        Property sColorNme As String
    End Class
End Class

Public Class Personal_Info
    Property sLastName As String
    Property sFrstName As String
    Property sMiddName As String
    Property sSuffixNm As String
    Property sMaidenNm As String
    Property cGenderCd As String
    Property dBirthDte As String
    Property sBirthPlc As String
    Property sHouseNox As String
    Property sAddressx As String
    Property sTownIDxx As String
    Property sMobileNo As String
    Property sEmailAdd As String
End Class

Public Class Financer_Info
    Property sLastName As String
    Property sFrstName As String
    Property sMiddName As String
    Property sSuffixNm As String
    Property sAddressx As String
    Property sCntryCde As String
    Property sMobileNo As String
    Property sWhatsApp As String
    Property sWChatApp As String
    Property sFbAccntx As String
    Property sEmailAdd As String
    Property sFIncomex As Decimal
    Property sReltionx As String
End Class

Public Class Product_Information
    Property sBrandIDx As String
    Property sModelIDx As String
    Property sColorIDx As String
    Property nSelPrice As Decimal
    Property dPricexxx As String
End Class

Public Class Payment_Information
    Property sTermIDxx As String
    Property nDownPaym As Decimal
End Class