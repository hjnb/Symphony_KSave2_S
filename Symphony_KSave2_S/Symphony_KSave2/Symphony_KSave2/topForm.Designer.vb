<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class topForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnSurveySlip = New System.Windows.Forms.Button()
        Me.btnMaster = New System.Windows.Forms.Button()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.rbtnPrint = New System.Windows.Forms.RadioButton()
        Me.btnTarget = New System.Windows.Forms.Button()
        Me.btnWriter = New System.Windows.Forms.Button()
        Me.btnReadOnly = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnSurveySlip
        '
        Me.btnSurveySlip.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnSurveySlip.Location = New System.Drawing.Point(26, 38)
        Me.btnSurveySlip.Name = "btnSurveySlip"
        Me.btnSurveySlip.Size = New System.Drawing.Size(228, 45)
        Me.btnSurveySlip.TabIndex = 0
        Me.btnSurveySlip.Text = "認　定　調　査　票 （登録用）"
        Me.btnSurveySlip.UseVisualStyleBackColor = True
        '
        'btnMaster
        '
        Me.btnMaster.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnMaster.Location = New System.Drawing.Point(26, 101)
        Me.btnMaster.Name = "btnMaster"
        Me.btnMaster.Size = New System.Drawing.Size(73, 60)
        Me.btnMaster.TabIndex = 1
        Me.btnMaster.Text = "マ ス タ"
        Me.btnMaster.UseVisualStyleBackColor = True
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Location = New System.Drawing.Point(144, 187)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(63, 16)
        Me.rbtnPreview.TabIndex = 2
        Me.rbtnPreview.Text = "ﾌﾟﾚﾋﾞｭｰ"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrint
        '
        Me.rbtnPrint.AutoSize = True
        Me.rbtnPrint.Location = New System.Drawing.Point(213, 187)
        Me.rbtnPrint.Name = "rbtnPrint"
        Me.rbtnPrint.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrint.TabIndex = 3
        Me.rbtnPrint.Text = "印刷"
        Me.rbtnPrint.UseVisualStyleBackColor = True
        '
        'btnTarget
        '
        Me.btnTarget.Location = New System.Drawing.Point(116, 101)
        Me.btnTarget.Name = "btnTarget"
        Me.btnTarget.Size = New System.Drawing.Size(59, 23)
        Me.btnTarget.TabIndex = 4
        Me.btnTarget.Text = "対象者"
        Me.btnTarget.UseVisualStyleBackColor = True
        '
        'btnWriter
        '
        Me.btnWriter.Location = New System.Drawing.Point(116, 138)
        Me.btnWriter.Name = "btnWriter"
        Me.btnWriter.Size = New System.Drawing.Size(59, 23)
        Me.btnWriter.TabIndex = 5
        Me.btnWriter.Text = "記入者"
        Me.btnWriter.UseVisualStyleBackColor = True
        Me.btnWriter.Visible = False
        '
        'btnReadOnly
        '
        Me.btnReadOnly.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnReadOnly.ForeColor = System.Drawing.Color.Blue
        Me.btnReadOnly.Location = New System.Drawing.Point(26, 229)
        Me.btnReadOnly.Name = "btnReadOnly"
        Me.btnReadOnly.Size = New System.Drawing.Size(228, 45)
        Me.btnReadOnly.TabIndex = 6
        Me.btnReadOnly.Text = "認　定　調　査　票 （閲覧用）"
        Me.btnReadOnly.UseVisualStyleBackColor = True
        '
        'topForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 302)
        Me.Controls.Add(Me.btnReadOnly)
        Me.Controls.Add(Me.btnWriter)
        Me.Controls.Add(Me.btnTarget)
        Me.Controls.Add(Me.rbtnPrint)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.btnMaster)
        Me.Controls.Add(Me.btnSurveySlip)
        Me.Name = "topForm"
        Me.Text = "認定調査票"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSurveySlip As System.Windows.Forms.Button
    Friend WithEvents btnMaster As System.Windows.Forms.Button
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrint As System.Windows.Forms.RadioButton
    Friend WithEvents btnTarget As System.Windows.Forms.Button
    Friend WithEvents btnWriter As System.Windows.Forms.Button
    Friend WithEvents btnReadOnly As System.Windows.Forms.Button

End Class
