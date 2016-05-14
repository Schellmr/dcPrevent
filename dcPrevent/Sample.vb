Imports System.Windows.Forms
Imports Gma.System.MouseKeyHook

Namespace dcPrevent
    Friend Class Sample
        Private m_GlobalHook As IKeyboardMouseEvents

        Public Sub Subscribe()
            ' Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents()

            AddHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
            AddHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
        End Sub

        Private Sub GlobalHookKeyPress(sender As Object, e As KeyPressEventArgs)
            Console.WriteLine("KeyPress: " & vbTab & "{0}", e.KeyChar)
        End Sub

        Private Sub GlobalHookMouseDownExt(sender As Object, e As MouseEventExtArgs)
            Console.WriteLine("MouseDown: " & vbTab & "{0}; " & vbTab & " System Timestamp: " & vbTab & "{1}", e.Button, e.Timestamp)

            ' uncommenting the following line will suppress the middle mouse button click
            ' if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        End Sub

        Public Sub Unsubscribe()
            RemoveHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
            RemoveHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress

            'It is recommened to dispose it
            m_GlobalHook.Dispose()
        End Sub
    End Class
End Namespace