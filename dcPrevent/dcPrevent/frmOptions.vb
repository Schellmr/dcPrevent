Imports System.ComponentModel
Imports System.Windows.Forms
Imports Gma.System.MouseKeyHook
Imports System.Threading
Imports Gma.System.MouseKeyHook.Implementation
Imports System.Runtime.InteropServices

Namespace dcPrevent
    Partial Public Class frmOptions
        Inherits Form

        Dim firstClick As Boolean = True
        Dim s As New Stopwatch
        Private m_Events As IKeyboardMouseEvents

        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOptions))
            Me.chkFilterDoubleClicks = New System.Windows.Forms.CheckBox()
            Me.lstLogs = New System.Windows.Forms.ListBox()
            Me.btnHide = New System.Windows.Forms.Button()
            Me.ntiShowForm = New System.Windows.Forms.NotifyIcon(Me.components)
            Me.SuspendLayout()
            '
            'chkFilterDoubleClicks
            '
            Me.chkFilterDoubleClicks.AutoSize = True
            Me.chkFilterDoubleClicks.Checked = True
            Me.chkFilterDoubleClicks.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkFilterDoubleClicks.Location = New System.Drawing.Point(13, 13)
            Me.chkFilterDoubleClicks.Name = "chkFilterDoubleClicks"
            Me.chkFilterDoubleClicks.Size = New System.Drawing.Size(116, 17)
            Me.chkFilterDoubleClicks.TabIndex = 0
            Me.chkFilterDoubleClicks.Text = "Filter Double Clicks"
            Me.chkFilterDoubleClicks.UseVisualStyleBackColor = True
            '
            'lstLogs
            '
            Me.lstLogs.FormattingEnabled = True
            Me.lstLogs.Location = New System.Drawing.Point(12, 37)
            Me.lstLogs.Name = "lstLogs"
            Me.lstLogs.Size = New System.Drawing.Size(278, 212)
            Me.lstLogs.TabIndex = 1
            '
            'btnHide
            '
            Me.btnHide.Location = New System.Drawing.Point(215, 7)
            Me.btnHide.Name = "btnHide"
            Me.btnHide.Size = New System.Drawing.Size(75, 23)
            Me.btnHide.TabIndex = 2
            Me.btnHide.Text = "Hide to Tray"
            Me.btnHide.UseVisualStyleBackColor = True
            '
            'ntiShowForm
            '
            Me.ntiShowForm.Icon = CType(resources.GetObject("ntiShowForm.Icon"), System.Drawing.Icon)
            Me.ntiShowForm.Text = "Show Form"
            Me.ntiShowForm.Visible = True
            '
            'frmOptions
            '
            Me.ClientSize = New System.Drawing.Size(302, 266)
            Me.Controls.Add(Me.btnHide)
            Me.Controls.Add(Me.lstLogs)
            Me.Controls.Add(Me.chkFilterDoubleClicks)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "frmOptions"
            Me.Text = "Options"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Public Sub New()
            InitializeComponent()
            SubscribeGlobal()
            AddHandler FormClosing, AddressOf Main_Closing
        End Sub
        Private Sub Main_Closing(sender As Object, e As CancelEventArgs)
            Unsubscribe()
        End Sub
        Private Sub SubscribeApplication()
            Unsubscribe()
            Subscribe(Hook.AppEvents())
        End Sub
        Private Sub SubscribeGlobal()
            Unsubscribe()
            Subscribe(Hook.GlobalEvents())
        End Sub

        Private Sub Subscribe(events As IKeyboardMouseEvents)
            m_Events = events

            s.Start()

            AddHandler m_Events.MouseDownExt, AddressOf HookManager_Supress
        End Sub
        Private Sub Unsubscribe()
            If m_Events Is Nothing Then
                Return
            End If


            RemoveHandler m_Events.MouseDownExt, AddressOf HookManager_Supress

            m_Events.Dispose()
            m_Events = Nothing
        End Sub

        Private Sub HookManager_Supress(sender As Object, e As MouseEventExtArgs)
            If e.Button = 1048576 Then
                If s.ElapsedMilliseconds < 50 And s.IsRunning And chkFilterDoubleClicks.Checked Then
                    e.Handled = True
                    Log(String.Format("Suppressed a DC" & vbTab & vbTab & " {0}" & vbLf, e.Button))
                    s.Reset()
                    s.Start()
                    Return
                End If
                Log(String.Format("MouseDown " & vbTab & vbTab & " {0}" & vbLf, e.Button))
                s.Reset()
                s.Start()
            End If
        End Sub

        Private Sub Log(text As String)
            If IsDisposed Then
                Return
            End If
            lstLogs.Items.Add(text)

            lstLogs.TopIndex = lstLogs.Items.Count - 1
        End Sub

        Friend WithEvents chkFilterDoubleClicks As CheckBox
        Friend WithEvents lstLogs As ListBox
        Friend WithEvents btnHide As Button
        Friend WithEvents ntiShowForm As NotifyIcon
        Private components As IContainer

        Private Sub btnHide_Click(sender As Object, e As EventArgs) Handles btnHide.Click
            MsgBox("Click the tray icon to get form back.")
            Hide()
        End Sub
        Private Sub ntiShowForm_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ntiShowForm.MouseDoubleClick
            Show()
        End Sub

    End Class
End Namespace

