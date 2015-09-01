<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.txtFormula = New System.Windows.Forms.TextBox()
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtFormula
        '
        Me.txtFormula.Location = New System.Drawing.Point(12, 12)
        Me.txtFormula.Name = "txtFormula"
        Me.txtFormula.Size = New System.Drawing.Size(273, 20)
        Me.txtFormula.TabIndex = 1
        Me.txtFormula.Text = "КОРЕНЬ(8;ОКРУГЛ(1,464*2))"
        '
        'btnCalc
        '
        Me.btnCalc.Location = New System.Drawing.Point(88, 38)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(120, 23)
        Me.btnCalc.TabIndex = 2
        Me.btnCalc.Text = "Рассчитать"
        Me.btnCalc.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(298, 75)
        Me.Controls.Add(Me.btnCalc)
        Me.Controls.Add(Me.txtFormula)
        Me.Name = "frmMain"
        Me.Text = "CalcFormulas"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtFormula As System.Windows.Forms.TextBox
    Friend WithEvents btnCalc As System.Windows.Forms.Button

End Class
