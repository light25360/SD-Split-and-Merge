'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   SD Split & Merge version 1.0
'   A simple software for split & Merge any file easily

'   Copyright (C) 2014  
'   Sajjad Tanha
'   Email : sajjadtanha20@gmail.com
'   Created: 20Dec14

'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http:'www.gnu.org/licenses/>.
'*****************************************************************************************

Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Threading

Public Class Form1
    Inherits System.Windows.Forms.Form
    Private Processing As Boolean
    Private backgroundThread As System.Threading.Thread
    Private NEWLINE As String = ControlChars.CrLf
    Private WithEvents _FileSplitMerge As New SplitMerge()
    Friend WithEvents cmdabout As System.Windows.Forms.Button
    Friend WithEvents cb1 As System.Windows.Forms.ComboBox
    Friend WithEvents FBD1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OFD1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtChunkSize As System.Windows.Forms.NumericUpDown
    Private _Progress As Double
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        InitializeComponent()

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents OptSplit As System.Windows.Forms.RadioButton
    Friend WithEvents OptMerge As System.Windows.Forms.RadioButton
    Friend WithEvents txtOutputFolder As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseFolder As System.Windows.Forms.Button
    Friend WithEvents chkOption As System.Windows.Forms.CheckBox
    Friend WithEvents lblChunkSize As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtChunkSize = New System.Windows.Forms.NumericUpDown()
        Me.cb1 = New System.Windows.Forms.ComboBox()
        Me.chkOption = New System.Windows.Forms.CheckBox()
        Me.cmdBrowseFolder = New System.Windows.Forms.Button()
        Me.lblChunkSize = New System.Windows.Forms.Label()
        Me.txtOutputFolder = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.cmdBrowseFile = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OptSplit = New System.Windows.Forms.RadioButton()
        Me.OptMerge = New System.Windows.Forms.RadioButton()
        Me.cmdabout = New System.Windows.Forms.Button()
        Me.FBD1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OFD1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1.SuspendLayout()
        CType(Me.txtChunkSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 176)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(392, 23)
        Me.ProgressBar1.TabIndex = 3
        Me.ProgressBar1.Visible = False
        '
        'cmdStart
        '
        Me.cmdStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdStart.Location = New System.Drawing.Point(59, 230)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(72, 32)
        Me.cmdStart.TabIndex = 9
        Me.cmdStart.Text = "&Start"
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Location = New System.Drawing.Point(131, 230)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 32)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "&Stop"
        '
        'cmdExit
        '
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.Location = New System.Drawing.Point(275, 230)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(72, 32)
        Me.cmdExit.TabIndex = 11
        Me.cmdExit.Text = "&Exit"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtChunkSize)
        Me.GroupBox1.Controls.Add(Me.cb1)
        Me.GroupBox1.Controls.Add(Me.chkOption)
        Me.GroupBox1.Controls.Add(Me.cmdBrowseFolder)
        Me.GroupBox1.Controls.Add(Me.lblChunkSize)
        Me.GroupBox1.Controls.Add(Me.txtOutputFolder)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtFileName)
        Me.GroupBox1.Controls.Add(Me.cmdBrowseFile)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.OptSplit)
        Me.GroupBox1.Controls.Add(Me.OptMerge)
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 208)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        '
        'txtChunkSize
        '
        Me.txtChunkSize.AutoSize = True
        Me.txtChunkSize.Location = New System.Drawing.Point(120, 97)
        Me.txtChunkSize.Maximum = New Decimal(New Integer() {1215752191, 23, 0, 0})
        Me.txtChunkSize.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtChunkSize.Name = "txtChunkSize"
        Me.txtChunkSize.Size = New System.Drawing.Size(101, 21)
        Me.txtChunkSize.TabIndex = 19
        Me.txtChunkSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtChunkSize.ThousandsSeparator = True
        Me.txtChunkSize.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'cb1
        '
        Me.cb1.FormattingEnabled = True
        Me.cb1.Items.AddRange(New Object() {"Byte", "Kilobyte", "Megabyte"})
        Me.cb1.Location = New System.Drawing.Point(236, 96)
        Me.cb1.Name = "cb1"
        Me.cb1.Size = New System.Drawing.Size(78, 21)
        Me.cb1.TabIndex = 18
        '
        'chkOption
        '
        Me.chkOption.AutoSize = True
        Me.chkOption.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.chkOption.Location = New System.Drawing.Point(120, 128)
        Me.chkOption.Name = "chkOption"
        Me.chkOption.Size = New System.Drawing.Size(12, 11)
        Me.chkOption.TabIndex = 17
        '
        'cmdBrowseFolder
        '
        Me.cmdBrowseFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdBrowseFolder.Location = New System.Drawing.Point(344, 72)
        Me.cmdBrowseFolder.Name = "cmdBrowseFolder"
        Me.cmdBrowseFolder.Size = New System.Drawing.Size(30, 20)
        Me.cmdBrowseFolder.TabIndex = 16
        Me.cmdBrowseFolder.Text = "..."
        '
        'lblChunkSize
        '
        Me.lblChunkSize.AutoSize = True
        Me.lblChunkSize.Location = New System.Drawing.Point(6, 98)
        Me.lblChunkSize.Name = "lblChunkSize"
        Me.lblChunkSize.Size = New System.Drawing.Size(67, 13)
        Me.lblChunkSize.TabIndex = 14
        Me.lblChunkSize.Text = "chunk  Size :"
        '
        'txtOutputFolder
        '
        Me.txtOutputFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOutputFolder.Location = New System.Drawing.Point(120, 72)
        Me.txtOutputFolder.Name = "txtOutputFolder"
        Me.txtOutputFolder.Size = New System.Drawing.Size(224, 21)
        Me.txtOutputFolder.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Save to :"
        '
        'txtFileName
        '
        Me.txtFileName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFileName.Location = New System.Drawing.Point(120, 48)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(224, 21)
        Me.txtFileName.TabIndex = 10
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdBrowseFile.Location = New System.Drawing.Point(344, 48)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(30, 20)
        Me.cmdBrowseFile.TabIndex = 9
        Me.cmdBrowseFile.Text = "..."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Filename :"
        '
        'OptSplit
        '
        Me.OptSplit.AutoSize = True
        Me.OptSplit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.OptSplit.Location = New System.Drawing.Point(24, 16)
        Me.OptSplit.Name = "OptSplit"
        Me.OptSplit.Size = New System.Drawing.Size(44, 17)
        Me.OptSplit.TabIndex = 7
        Me.OptSplit.Text = "Split"
        '
        'OptMerge
        '
        Me.OptMerge.AutoSize = True
        Me.OptMerge.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.OptMerge.Location = New System.Drawing.Point(96, 16)
        Me.OptMerge.Name = "OptMerge"
        Me.OptMerge.Size = New System.Drawing.Size(54, 17)
        Me.OptMerge.TabIndex = 6
        Me.OptMerge.Text = "Merge"
        '
        'cmdabout
        '
        Me.cmdabout.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdabout.Location = New System.Drawing.Point(203, 230)
        Me.cmdabout.Name = "cmdabout"
        Me.cmdabout.Size = New System.Drawing.Size(72, 32)
        Me.cmdabout.TabIndex = 15
        Me.cmdabout.Text = "&About"
        Me.cmdabout.UseVisualStyleBackColor = True
        '
        'OFD1
        '
        Me.OFD1.FileName = "OpenFileDialog1"
        Me.OFD1.Title = "انتخاب فایل"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(418, 271)
        Me.Controls.Add(Me.cmdabout)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdStart)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SD Split & Merge"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.txtChunkSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        ProgressBar1.Value = 0
        If ValidateData() Then
            _Progress = 0
            Processing = True
            cmdCancel.Enabled = True
            GroupBox1.Enabled = False
            ProgressBar1.Visible = True
            If OptSplit.Checked Then

                Me.SplitFile()
            Else
                Me.MergeFiles()
            End If
        End If
    End Sub
    ''''
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        If Processing Then
            ShowMessage()
        Else
            Me.Close()
        End If
    End Sub
    '''
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If backgroundThread.IsAlive Then backgroundThread.Abort()
        Processing = False
        GroupBox1.Enabled = True
    End Sub
    '''
    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        If OptSplit.Checked = True Then
            OFD1.Filter = "All Files (*.*)|*.*"
            OFD1.FileName = Nothing
        ElseIf OptMerge.Checked = True Then
            OFD1.Filter = "001 (*.001)|*.001"
            OFD1.FileName = "001 File"
        End If
        If OFD1.ShowDialog() = DialogResult.OK Then
            txtFileName.Text = OFD1.FileName
            OFD1.Dispose()
            Dim sfilesize As Long = My.Computer.FileSystem.GetFileInfo(txtFileName.Text).Length
            If sfilesize > 1024 Then
                txtChunkSize.Value = sfilesize / 10
            End If
        End If
    End Sub
    '''
    Private Sub cmdBrowseFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFolder.Click
        FBD1.Description = "Choose output folder "
        If FBD1.ShowDialog(Me) = DialogResult.OK Then
            txtOutputFolder.Text = FBD1.SelectedPath
            FBD1.Dispose()
        End If
    End Sub
    '''
    Private Sub ShowMessage()
        If OptSplit.Checked Then
            MsgBox("Split process is running" & ControlChars.CrLf & "Please wait until process finish " & ControlChars.CrLf & "Or to stop the process, click on stop button.")
        Else
            MsgBox("Merge process is running" & ControlChars.CrLf & "Please wait until process finish " & ControlChars.CrLf & "Or to stop the process, click on stop button.")
        End If
    End Sub
    '''
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Processing Then
            ShowMessage()
            e.Cancel = True
        End If
    End Sub
    '''
    Private Function ValidateData() As Boolean
        Dim _error As String
        Dim _FileSize As Long
        _error = ""
        If txtFileName.Text = "" Then
            _error += "Filename should not be empty." & NEWLINE
        Else
            If Not System.IO.File.Exists(txtFileName.Text) Then
                _error += txtFileName.Text & " Not found." & NEWLINE
            Else
                _FileSize = _FileSplitMerge.FileSize(txtFileName.Text)
            End If
        End If
        If OptSplit.Checked Then
            If txtChunkSize.Value >= _FileSize Then
                _error += "Chunk size is bigger than original file size." & NEWLINE
            End If

        End If
        If _error = "" Then
            Return True
        Else
            MsgBox(_error, MsgBoxStyle.Information, Me.Text)
            Return False
        End If
    End Function
    '''
    Private Sub SplitFile()
        Dim tcs As Long = 0
        If cb1.SelectedItem = cb1.Items(0) Then
            tcs = txtChunkSize.Value
        ElseIf cb1.SelectedItem = cb1.Items(1) Then
            tcs = txtChunkSize.Value * 1024
        Else
            tcs = (txtChunkSize.Value * 1024) * 1024
        End If
        With _FileSplitMerge
            .ChunkSize = tcs
            .FileName = txtFileName.Text
            .OutputPath = txtOutputFolder.Text
            .DeleteFileAfterSplit = chkOption.Checked
            backgroundThread = New Threading.Thread(AddressOf .SplitFile)
            backgroundThread.Start()
        End With
    End Sub
    '''
    Private Sub MergeFiles()
        With _FileSplitMerge
            .FileName = txtFileName.Text
            .OutputPath = txtOutputFolder.Text
            .DeleteFilesAfterMerge = chkOption.Checked
            backgroundThread = New Threading.Thread(AddressOf .MergeFile)
            backgroundThread.Start()
        End With
    End Sub
    '''
    Private Sub _FileSplitMerge_FileMergeCompleted(ByVal ErrorMessage As String) Handles _FileSplitMerge.FileMergeCompleted
        If ErrorMessage = "" Then
            MsgBox("Files successfully merged")
        Else
            MsgBox("An error occurred during merge process." & NEWLINE & ErrorMessage)
        End If
        ThreadPool.QueueUserWorkItem(AddressOf Me.ResetControls)
    End Sub
    '''
    Private Sub _FileSplitMerge_FileSplitCompleted(ByVal ErrorMessage As String) Handles _FileSplitMerge.FileSplitCompleted
        If ErrorMessage = "" Then
            MsgBox("File successfully been split")
        Else
            MsgBox("An error occurred during split process." & NEWLINE & ErrorMessage)
        End If
        ThreadPool.QueueUserWorkItem(AddressOf Me.ResetControls)
    End Sub
    '''
    Private Sub _FileSplitMerge_UpdateProgress(ByVal Progress As Double) Handles _FileSplitMerge.UpdateProgress
        _Progress += Progress
        upproinvoke(_Progress)
    End Sub
    '''
    Private Sub upproinvoke(ByVal value As Long)
        If ProgressBar1.InvokeRequired Then
            If ProgressBar1.Value < ProgressBar1.Maximum Then
                ProgressBar1.Invoke(New Action(Of Long)(AddressOf upproinvoke), value)
            End If
        Else
            If ProgressBar1.Value < ProgressBar1.Maximum Then
                ProgressBar1.Value = value
            End If
        End If
    End Sub
    Private Sub OptSplit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSplit.CheckedChanged
        If OptSplit.Checked Then
            chkOption.Text = "Delete file after split"
            lblChunkSize.Visible = True
            txtChunkSize.Visible = True
            cb1.Visible = True
        Else
            chkOption.Text = "Delete files after merge"
            lblChunkSize.Visible = False
            txtChunkSize.Visible = False
            cb1.Visible = False
        End If
    End Sub
    '''
    Private Sub OptMerge_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptMerge.CheckedChanged
        If OptSplit.Checked Then
            chkOption.Text = "Delete file after split"
            lblChunkSize.Visible = True
            txtChunkSize.Visible = True
            cb1.Visible = True
        Else
            chkOption.Text = "Delete files after merge"
            lblChunkSize.Visible = False
            txtChunkSize.Visible = False
            cb1.Visible = False
        End If
    End Sub
    '''
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OptSplit.Checked = True
        cmdCancel.Enabled = False
        chkOption.Text = "Delete file after split"
        cb1.SelectedItem = cb1.Items(0)
    End Sub
    '''
    Private Sub ResetControls()
        Processing = False
        AccessControl()
    End Sub
    Private Sub AccessControl()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf AccessControl))
        Else
           Me.cmdCancel.Enabled = False
            Me.GroupBox1.Enabled = True
            Me.ProgressBar1.Visible = False
        End If
    End Sub
    Private Sub btnabout_Click(sender As Object, e As EventArgs) Handles cmdabout.Click
        AboutBox1.ShowDialog()
    End Sub
End Class
