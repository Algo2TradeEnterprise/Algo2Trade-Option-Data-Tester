<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.cmbSymbol = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblMainProgress = New System.Windows.Forms.Label()
        Me.txtAvrgSprdPrcntg = New System.Windows.Forms.TextBox()
        Me.lblAvrgSprdPrcntg = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(25, 308)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(695, 98)
        Me.lblProgress.TabIndex = 0
        Me.lblProgress.Text = "lblProgress"
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(547, 70)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(84, 43)
        Me.btnStart.TabIndex = 1
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(647, 70)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(84, 43)
        Me.btnStop.TabIndex = 5
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'lblFilePath
        '
        Me.lblFilePath.AutoSize = True
        Me.lblFilePath.Location = New System.Drawing.Point(25, 83)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.Size = New System.Drawing.Size(126, 17)
        Me.lblFilePath.TabIndex = 6
        Me.lblFilePath.Text = "Template File Path"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(434, 81)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(41, 23)
        Me.btnBrowse.TabIndex = 8
        Me.btnBrowse.Text = "..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(160, 82)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(268, 22)
        Me.txtFilePath.TabIndex = 7
        '
        'cmbSymbol
        '
        Me.cmbSymbol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSymbol.FormattingEnabled = True
        Me.cmbSymbol.Location = New System.Drawing.Point(115, 17)
        Me.cmbSymbol.Name = "cmbSymbol"
        Me.cmbSymbol.Size = New System.Drawing.Size(320, 24)
        Me.cmbSymbol.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Symbol List:"
        '
        'lblMainProgress
        '
        Me.lblMainProgress.Location = New System.Drawing.Point(25, 281)
        Me.lblMainProgress.Name = "lblMainProgress"
        Me.lblMainProgress.Size = New System.Drawing.Size(695, 27)
        Me.lblMainProgress.TabIndex = 10
        '
        'txtAvrgSprdPrcntg
        '
        Me.txtAvrgSprdPrcntg.Location = New System.Drawing.Point(604, 18)
        Me.txtAvrgSprdPrcntg.Name = "txtAvrgSprdPrcntg"
        Me.txtAvrgSprdPrcntg.Size = New System.Drawing.Size(69, 22)
        Me.txtAvrgSprdPrcntg.TabIndex = 12
        Me.txtAvrgSprdPrcntg.Text = "1"
        '
        'lblAvrgSprdPrcntg
        '
        Me.lblAvrgSprdPrcntg.AutoSize = True
        Me.lblAvrgSprdPrcntg.Location = New System.Drawing.Point(469, 19)
        Me.lblAvrgSprdPrcntg.Name = "lblAvrgSprdPrcntg"
        Me.lblAvrgSprdPrcntg.Size = New System.Drawing.Size(127, 17)
        Me.lblAvrgSprdPrcntg.TabIndex = 11
        Me.lblAvrgSprdPrcntg.Text = "Average Spread %"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(750, 418)
        Me.Controls.Add(Me.txtAvrgSprdPrcntg)
        Me.Controls.Add(Me.lblAvrgSprdPrcntg)
        Me.Controls.Add(Me.lblMainProgress)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbSymbol)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.lblFilePath)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.lblProgress)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Algo2Trade Option Data Tester"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblProgress As Label
    Friend WithEvents btnStart As Button
    Friend WithEvents btnStop As Button
    Friend WithEvents lblFilePath As Label
    Friend WithEvents btnBrowse As Button
    Friend WithEvents OpenFileDialog As OpenFileDialog
    Friend WithEvents txtFilePath As TextBox
    Friend WithEvents cmbSymbol As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblMainProgress As Label
    Friend WithEvents txtAvrgSprdPrcntg As TextBox
    Friend WithEvents lblAvrgSprdPrcntg As Label
End Class
