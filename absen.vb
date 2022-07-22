Imports System.Data.OleDb

Public Class Absen
    Dim tgl1 As DateTime
    Dim tgl2 As DateTime
    Dim wkt As DateTime
    Dim sel As TimeSpan

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dashboard.Enabled = False
        Me.Text = "Form Absensi"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.TextBox1.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tgl1 = Format(Now, "MM/dd/yyyy")
        tgl2 = DateTime.Now
        Dim simpan As String
        Dim who As String
        If TextBox1.Text = "" Then
            MsgBox("Anda belum memasukan NIK")
            TextBox1.Clear()
            TextBox1.Focus()
        Else
            Call konek.Koneksi()
            cmd = New OleDbCommand("SELECT * FROM masuk WHERE nik = '" & TextBox1.Text & "' AND tgl = Date()", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                    cmd = New OleDbCommand("Select * FROM masuk INNER JOIN pulang ON masuk.id_masuk = pulang.id_masuk WHERE nik = '" & TextBox1.Text & "'", Conn)
                    dr = cmd.ExecuteReader
                    dr.Read()
                    Select Case dr.HasRows
                        Case True
                            MsgBox("Anda sudah mengabsen pulang", vbInformation)
                            TextBox1.Clear()
                            TextBox1.Focus()
                    Case False
                        cmd = New OleDbCommand("Select * from masuk where nik = '" & TextBox1.Text & "'", Conn)
                        dr = cmd.ExecuteReader
                        dr.Read()
                        Dim jm As Integer = dr.Item("id_masuk")
                        wkt = dr.Item("jam_msk")
                        sel = tgl2 - wkt
                        TextBox2.Text = sel.ToString
                        simpan = "INSERT INTO pulang (id_masuk,jam_krj) VALUES ('" & jm & "','" & TextBox2.Text & "')"
                        jalankansql(simpan)
                        cmd = New OleDbCommand("Select nama_lkp from data_karyawan where nik = '" & TextBox1.Text & "'", Conn)
                        dr = cmd.ExecuteReader
                        dr.Read()
                        who = dr.Item("nama_lkp")
                        MsgBox("Terimakasih, Sampai jumpa " & who, vbInformation)
                        TextBox1.Clear()
                        TextBox2.Clear()
                        TextBox1.Focus()
                End Select
            Else
                Try
                    cmd = New OleDbCommand("Select nik from data_karyawan where nik = '" & TextBox1.Text & "'", Conn)
                    dr = cmd.ExecuteReader
                    dr.Read()
                    If dr.HasRows Then
                        simpan = "INSERT INTO masuk (nik,tgl,jam_msk) VALUES ('" & TextBox1.Text & "','" & tgl1 & "','" & tgl2 & "')"
                        jalankansql(simpan)
                        Call konek.Koneksi()
                        cmd = New OleDbCommand("Select nama_lkp from data_karyawan where nik = '" & TextBox1.Text & "'", Conn)
                        dr = cmd.ExecuteReader
                        dr.Read()
                        who = dr.Item("nama_lkp")
                        MsgBox("Selamat datang " & who, vbInformation)
                        TextBox1.Clear()
                        TextBox1.Focus()
                    Else
                        MsgBox("Anda tidak terdaftar sebagai karyawan", vbInformation)
                        TextBox1.Clear()
                        TextBox1.Focus()
                    End If
                Catch ex As Exception
                    MsgBox("Data error hubungi admin", vbInformation)
                    Exit Sub
                End Try
            End If
            Conn.Close()
        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs)
        tgl2 = DateTime.Now
        sel = tgl2 - tgl1
        TextBox2.Text = sel.ToString
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Label1.Text = Format(Now, "dd-MM-yyyy hh:mm:ss")
        Dim d = TimeSerial(18, 0, 0)
        If TimeOfDay >= d Then
            Me.TextBox1.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        dashboard.Show()
        dashboard.Enabled = True
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then Button1.Focus()
    End Sub
End Class
