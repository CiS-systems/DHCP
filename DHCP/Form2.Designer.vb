<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
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
        Me.dgv4 = New System.Windows.Forms.DataGridView()
        Me.MacGet = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ListGet = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescGet = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.dgv4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgv4
        '
        Me.dgv4.AllowUserToAddRows = False
        Me.dgv4.AllowUserToDeleteRows = False
        Me.dgv4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv4.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MacGet, Me.ListGet, Me.DescGet, Me.Column1, Me.Column2, Me.Column3})
        Me.dgv4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv4.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dgv4.Location = New System.Drawing.Point(0, 0)
        Me.dgv4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dgv4.MultiSelect = False
        Me.dgv4.Name = "dgv4"
        Me.dgv4.RowHeadersVisible = False
        Me.dgv4.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgv4.Size = New System.Drawing.Size(1113, 390)
        Me.dgv4.TabIndex = 16
        '
        'MacGet
        '
        Me.MacGet.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.MacGet.HeaderText = "MAC"
        Me.MacGet.Name = "MacGet"
        Me.MacGet.ReadOnly = True
        '
        'ListGet
        '
        Me.ListGet.HeaderText = "Status"
        Me.ListGet.Name = "ListGet"
        Me.ListGet.ReadOnly = True
        Me.ListGet.Width = 80
        '
        'DescGet
        '
        Me.DescGet.HeaderText = "Beschreibung"
        Me.DescGet.Name = "DescGet"
        Me.DescGet.ReadOnly = True
        Me.DescGet.Width = 200
        '
        'Column1
        '
        Me.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column1.HeaderText = "Neue MAC"
        Me.Column1.Name = "Column1"
        '
        'Column2
        '
        Me.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.Column2.HeaderText = "Neue Status"
        Me.Column2.Items.AddRange(New Object() {"Allow", "Deny"})
        Me.Column2.Name = "Column2"
        Me.Column2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column2.Width = 80
        '
        'Column3
        '
        Me.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Column3.HeaderText = "Neue Beschreibung"
        Me.Column3.Name = "Column3"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Button2)
        Me.Panel4.Controls.Add(Me.Button1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(0, 390)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1113, 60)
        Me.Panel4.TabIndex = 18
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(855, 13)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 35)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Abbrechen"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(988, 11)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 35)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Ändern"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1113, 450)
        Me.Controls.Add(Me.dgv4)
        Me.Controls.Add(Me.Panel4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Änderungen"
        CType(Me.dgv4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgv4 As DataGridView
    Friend WithEvents MacGet As DataGridViewTextBoxColumn
    Friend WithEvents ListGet As DataGridViewTextBoxColumn
    Friend WithEvents DescGet As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewComboBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
End Class
