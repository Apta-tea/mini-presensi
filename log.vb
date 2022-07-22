Public Class log

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.


    Public ghn As Integer
    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'dashboard.Enabled = False
        Me.UsernameTextBox.Focus()
    End Sub

    Private Sub UsernameTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles UsernameTextBox.KeyPress
        If e.KeyChar = Chr(13) Then PasswordTextBox.Focus()
    End Sub

    Private Sub PasswordTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles PasswordTextBox.KeyPress
        If e.KeyChar = Chr(13) Then Button2.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'dashboard.Enabled = True
        'conf = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call konek.Koneksi()
        cmd = New OleDb.OleDbCommand("SELECT * FROM sandi WHERE nama='" & UsernameTextBox.Text & "' AND sandi='" & PasswordTextBox.Text & "'", Conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Me.Close()
            admin.Show()
        Else
            MsgBox("Password/Username salah!", vbInformation)
            UsernameTextBox.Focus()
        End If
    End Sub
End Class
