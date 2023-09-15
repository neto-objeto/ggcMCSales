<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutboundCall
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutboundCall))
        Me.gbxPanel01 = New System.Windows.Forms.GroupBox()
        Me.rbtTypex02 = New System.Windows.Forms.RadioButton()
        Me.rbtTypex01 = New System.Windows.Forms.RadioButton()
        Me.gbxPanel03 = New System.Windows.Forms.GroupBox()
        Me.chkInclude00 = New System.Windows.Forms.CheckBox()
        Me.chkInclude02 = New System.Windows.Forms.CheckBox()
        Me.chkInclude01 = New System.Windows.Forms.CheckBox()
        Me.cmdButton01 = New System.Windows.Forms.Button()
        Me.cmdButton00 = New System.Windows.Forms.Button()
        Me.gbxPanel04 = New System.Windows.Forms.GroupBox()
        Me.txtField02 = New System.Windows.Forms.TextBox()
        Me.txtField01 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbxPanel01.SuspendLayout()
        Me.gbxPanel03.SuspendLayout()
        Me.gbxPanel04.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbxPanel01
        '
        Me.gbxPanel01.Controls.Add(Me.rbtTypex02)
        Me.gbxPanel01.Controls.Add(Me.rbtTypex01)
        Me.gbxPanel01.Location = New System.Drawing.Point(12, 13)
        Me.gbxPanel01.Name = "gbxPanel01"
        Me.gbxPanel01.Size = New System.Drawing.Size(200, 88)
        Me.gbxPanel01.TabIndex = 0
        Me.gbxPanel01.TabStop = False
        Me.gbxPanel01.Text = "Report Type"
        '
        'rbtTypex02
        '
        Me.rbtTypex02.AutoSize = True
        Me.rbtTypex02.Checked = True
        Me.rbtTypex02.Location = New System.Drawing.Point(16, 43)
        Me.rbtTypex02.Name = "rbtTypex02"
        Me.rbtTypex02.Size = New System.Drawing.Size(52, 17)
        Me.rbtTypex02.TabIndex = 1
        Me.rbtTypex02.TabStop = True
        Me.rbtTypex02.Text = "Detail"
        Me.rbtTypex02.UseVisualStyleBackColor = True
        '
        'rbtTypex01
        '
        Me.rbtTypex01.AutoSize = True
        Me.rbtTypex01.Location = New System.Drawing.Point(16, 20)
        Me.rbtTypex01.Name = "rbtTypex01"
        Me.rbtTypex01.Size = New System.Drawing.Size(68, 17)
        Me.rbtTypex01.TabIndex = 0
        Me.rbtTypex01.Text = "Summary"
        Me.rbtTypex01.UseVisualStyleBackColor = True
        '
        'gbxPanel03
        '
        Me.gbxPanel03.Controls.Add(Me.chkInclude00)
        Me.gbxPanel03.Controls.Add(Me.chkInclude02)
        Me.gbxPanel03.Controls.Add(Me.chkInclude01)
        Me.gbxPanel03.Location = New System.Drawing.Point(218, 13)
        Me.gbxPanel03.Name = "gbxPanel03"
        Me.gbxPanel03.Size = New System.Drawing.Size(200, 88)
        Me.gbxPanel03.TabIndex = 2
        Me.gbxPanel03.TabStop = False
        Me.gbxPanel03.Text = "Category"
        '
        'chkInclude00
        '
        Me.chkInclude00.AutoSize = True
        Me.chkInclude00.Location = New System.Drawing.Point(16, 21)
        Me.chkInclude00.Name = "chkInclude00"
        Me.chkInclude00.Size = New System.Drawing.Size(37, 17)
        Me.chkInclude00.TabIndex = 4
        Me.chkInclude00.Text = "All"
        Me.chkInclude00.UseVisualStyleBackColor = True
        '
        'chkInclude02
        '
        Me.chkInclude02.AutoSize = True
        Me.chkInclude02.Location = New System.Drawing.Point(16, 67)
        Me.chkInclude02.Name = "chkInclude02"
        Me.chkInclude02.Size = New System.Drawing.Size(72, 17)
        Me.chkInclude02.TabIndex = 1
        Me.chkInclude02.Text = "Financing"
        Me.chkInclude02.UseVisualStyleBackColor = True
        '
        'chkInclude01
        '
        Me.chkInclude01.AutoSize = True
        Me.chkInclude01.Location = New System.Drawing.Point(16, 44)
        Me.chkInclude01.Name = "chkInclude01"
        Me.chkInclude01.Size = New System.Drawing.Size(60, 17)
        Me.chkInclude01.TabIndex = 0
        Me.chkInclude01.Text = "Default"
        Me.chkInclude01.UseVisualStyleBackColor = True
        '
        'cmdButton01
        '
        Me.cmdButton01.Image = CType(resources.GetObject("cmdButton01.Image"), System.Drawing.Image)
        Me.cmdButton01.Location = New System.Drawing.Point(439, 20)
        Me.cmdButton01.Name = "cmdButton01"
        Me.cmdButton01.Size = New System.Drawing.Size(97, 40)
        Me.cmdButton01.TabIndex = 6
        Me.cmdButton01.Text = "&Ok"
        Me.cmdButton01.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButton01.UseVisualStyleBackColor = True
        '
        'cmdButton00
        '
        Me.cmdButton00.Image = CType(resources.GetObject("cmdButton00.Image"), System.Drawing.Image)
        Me.cmdButton00.Location = New System.Drawing.Point(439, 60)
        Me.cmdButton00.Name = "cmdButton00"
        Me.cmdButton00.Size = New System.Drawing.Size(97, 40)
        Me.cmdButton00.TabIndex = 7
        Me.cmdButton00.Text = "&Cancel"
        Me.cmdButton00.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButton00.UseVisualStyleBackColor = True
        '
        'gbxPanel04
        '
        Me.gbxPanel04.Controls.Add(Me.txtField02)
        Me.gbxPanel04.Controls.Add(Me.txtField01)
        Me.gbxPanel04.Controls.Add(Me.Label2)
        Me.gbxPanel04.Controls.Add(Me.Label1)
        Me.gbxPanel04.Location = New System.Drawing.Point(12, 103)
        Me.gbxPanel04.Name = "gbxPanel04"
        Me.gbxPanel04.Size = New System.Drawing.Size(406, 75)
        Me.gbxPanel04.TabIndex = 3
        Me.gbxPanel04.TabStop = False
        Me.gbxPanel04.Text = "Range"
        '
        'txtField02
        '
        Me.txtField02.Location = New System.Drawing.Point(64, 44)
        Me.txtField02.Name = "txtField02"
        Me.txtField02.Size = New System.Drawing.Size(147, 20)
        Me.txtField02.TabIndex = 3
        '
        'txtField01
        '
        Me.txtField01.Location = New System.Drawing.Point(64, 20)
        Me.txtField01.Name = "txtField01"
        Me.txtField01.Size = New System.Drawing.Size(147, 20)
        Me.txtField01.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Thru"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'frmCRRequest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(548, 186)
        Me.Controls.Add(Me.gbxPanel04)
        Me.Controls.Add(Me.cmdButton01)
        Me.Controls.Add(Me.cmdButton00)
        Me.Controls.Add(Me.gbxPanel03)
        Me.Controls.Add(Me.gbxPanel01)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmCRRequest"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Outbound Call Reports"
        Me.gbxPanel01.ResumeLayout(False)
        Me.gbxPanel01.PerformLayout()
        Me.gbxPanel03.ResumeLayout(False)
        Me.gbxPanel03.PerformLayout()
        Me.gbxPanel04.ResumeLayout(False)
        Me.gbxPanel04.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbxPanel01 As System.Windows.Forms.GroupBox
    Friend WithEvents rbtTypex02 As System.Windows.Forms.RadioButton
    Friend WithEvents rbtTypex01 As System.Windows.Forms.RadioButton
    Friend WithEvents gbxPanel03 As System.Windows.Forms.GroupBox
    Friend WithEvents chkInclude02 As System.Windows.Forms.CheckBox
    Friend WithEvents chkInclude01 As System.Windows.Forms.CheckBox
    Friend WithEvents cmdButton01 As System.Windows.Forms.Button
    Friend WithEvents cmdButton00 As System.Windows.Forms.Button
    Friend WithEvents gbxPanel04 As System.Windows.Forms.GroupBox
    Friend WithEvents txtField02 As System.Windows.Forms.TextBox
    Friend WithEvents txtField01 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkInclude00 As System.Windows.Forms.CheckBox
End Class
