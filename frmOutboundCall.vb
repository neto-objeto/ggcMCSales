
Public Class frmOutboundCall

    Private pb_ChkdOk As Boolean
    Private pn_Loaded As Integer
    Private p_oDriver As ggcAppDriver.GRider

    Public WriteOnly Property GRider() As ggcAppDriver.GRider
        Set(ByVal foValue As ggcAppDriver.GRider)
            p_oDriver = foValue
        End Set
    End Property

    Public Function isOkey() As Boolean
        Return pb_ChkdOk
    End Function

    Private Sub chkInclude00_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInclude00.CheckedChanged
        If chkInclude00.Checked Then
            chkInclude01.Checked = True
            chkInclude02.Checked = True

            chkInclude01.Enabled = False
            chkInclude02.Enabled = False
        Else
            chkInclude01.Enabled = True
            chkInclude02.Enabled = True
        End If
    End Sub

    Private Sub cmdButton01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton01.Click

        If Not (chkInclude00.Checked Or
                chkInclude01.Checked Or
                chkInclude02.Checked) Then
            MsgBox("There are items selected in the INCLUDE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub

        ElseIf Not (IsDate(txtField01.Text) And
                    IsDate(txtField02.Text)) Then

            MsgBox("There are invalid date in the RANGE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub

        ElseIf CDate(txtField01.Text) > CDate(txtField02.Text) Then
            MsgBox("FROM parameter seems to be higher than THRU in the RANGE group." & vbCrLf &
                   "Please check your entry try again!", vbOKOnly, "Parameter Validation")
            Exit Sub
        End If

        pb_ChkdOk = True

        Me.Hide()
    End Sub

    Private Sub cmdButton00_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton00.Click
        pb_ChkdOk = False
        Me.Hide()
    End Sub

    Private Sub txtField01_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtField01.Validated
        If IsDate(txtField01.Text) Then
            txtField01.Text = Format(CDate(txtField01.Text), "yyyy-MM-dd")
        Else
            txtField01.Text = Format(Now(), "yyyy-MM-dd")
        End If
    End Sub

    Private Sub txtField02_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtField02.Validated
        If IsDate(txtField02.Text) Then
            txtField02.Text = Format(CDate(txtField02.Text), "yyyy-MM-dd")
        Else
            txtField02.Text = Format(Now(), "yyyy-MM-dd")
        End If
    End Sub
End Class
