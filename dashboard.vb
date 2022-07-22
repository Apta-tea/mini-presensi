Public Class dashboard
    Public konf As Integer = 0
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Aplikasi Presensi Ver-1.0"
    End Sub

    Private Sub HapusDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HapusDataToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub InputEditDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InputEditDataToolStripMenuItem.Click
        Absen.Show()
    End Sub
    Private Sub MenuAdminToolStripMenuItem_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuAdminToolStripMenuItem.DropDownItemClicked
        log.Show()
    End Sub

    Private Sub TentangToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TentangToolStripMenuItem1.Click
        AboutBox1.Show()
    End Sub
End Class
