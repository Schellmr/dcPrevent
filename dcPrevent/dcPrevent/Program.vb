Imports System.Windows.Forms

Namespace dcPrevent
    Friend NotInheritable Class Program
        Private Sub New()
        End Sub
        ''' <summary>
        '''     The main entry point for the application.
        ''' </summary>
        <STAThread>
        Private Shared Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New frmOptions())
        End Sub
    End Class
End Namespace
