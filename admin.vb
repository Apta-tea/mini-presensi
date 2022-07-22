Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel
Public Class admin
    Private Sub admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dashboard.Enabled = False
        dgv1()
        dgv3()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        dashboard.Enabled = True
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call konek.Koneksi()
            cmd = New OleDbCommand("SELECT * FROM data_karyawan WHERE nik = '" & TextBox1.Text & "'", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                TextBox2.Text = dr.Item("nama_lkp")
                TextBox3.Text = dr.Item("alamat")
                TextBox4.Text = dr.Item("no_tlp")
                Conn.Close()
            Else
                Conn.Close()
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then TextBox3.Focus()
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Chr(13) Then TextBox4.Focus()
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = Chr(13) Then Button2.Focus()
    End Sub
    Sub dgv1()
        konek.Koneksi()
        da = New OleDbDataAdapter("Select * from data_karyawan", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "data_karyawan")
        DataGridView1.DataSource = (ds.Tables("data_karyawan"))
        DataGridView1.Enabled = True
        Conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        bersih1()
        dgv1()
        TextBox1.Focus()
    End Sub
    Sub bersih1()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim simpan As String
        konek.Koneksi()
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Field kosong terdeteksi!", vbExclamation, "Perhatian")
        Else
            cmd = New OleDbCommand("Select nik from data_karyawan where nik = '" & TextBox1.Text & "'", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                simpan = "UPDATE data_karyawan SET nama_lkp= '" & TextBox2.Text & "', alamat ='" & TextBox3.Text & "', no_tlp = '" & TextBox4.Text & "' WHERE nik = '" & TextBox1.Text & "'"
                konek.jalankansql(simpan)
                dgv1()
                DataGridView1.Refresh()
                Me.Cursor = Cursors.Default
                MsgBox("Data telah diupdate!", vbInformation)
                bersih1()
                TextBox1.Focus()
            Else
                simpan = "INSERT INTO data_karyawan(nik,nama_lkp,alamat,no_tlp) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                konek.jalankansql(simpan)
                dgv1()
                DataGridView1.Refresh()
                MsgBox("Data telah disimpan!", vbInformation)
                bersih1()
                TextBox1.Focus()
            End If
        End If
        Conn.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim hapus As String
        Dim pesan As Integer
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Field kosong terdeteksi!", vbExclamation, "Peringatan")
        Else
            pesan = MsgBox("Apakah anda yakin akan menghapus Data pada Database .. " + TextBox2.Text, vbExclamation + vbYesNo, "Perhatian")
            If pesan = vbNo Then Exit Sub
            hapus = "DELETE FROM data_karyawan WHERE nik='" & TextBox1.Text & "'"
            konek.jalankansql(hapus)
            dgv1()
            DataGridView1.Refresh()
            MsgBox("Data telah dihapus!", vbInformation)
            TextBox1.Focus()
        End If
    End Sub
    Private Sub isiTextBox(ByVal x As Integer)
        Try
            TextBox1.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox2.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(2).Value
            TextBox4.Text = DataGridView1.Rows(x).Cells(3).Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        isiTextBox(e.RowIndex)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Close()
        dashboard.Enabled = True
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Me.Close()
        dashboard.Enabled = True
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim pesan As Integer
        Dim simpan As String
        If TextBox8.Text = "" Or TextBox9.Text = "" Then
            MsgBox("Field kosong terdeteksi", vbExclamation, "Perhatian")
        Else
            konek.Koneksi()
            cmd = New OleDbCommand("SELECT * FROM sandi WHERE nama='" & TextBox8.Text & "'", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                pesan = MsgBox("Anda yakin akan mengupdate sandi " + TextBox8.Text, vbExclamation + vbYesNo, "Peringatan")
                If pesan = vbNo Then
                    Exit Sub
                End If
                simpan = "UPDATE sandi SET sandi='" & TextBox9.Text & "' WHERE nama='" & TextBox8.Text & "'"
                konek.jalankansql(simpan)
                MsgBox("Data telah diupdate", vbExclamation)
                dgv3()
                DataGridView3.Refresh()
                TextBox8.Focus()
                bersih3()
            Else
                simpan = "INSERT INTO sandi (nama,sandi) VALUES ('" & TextBox8.Text & "','" & TextBox9.Text & "')"
                konek.jalankansql(simpan)
                MsgBox("Data telah disimpan", vbExclamation)
                dgv3()
                DataGridView3.Refresh()
                TextBox8.Focus()
                bersih3()
            End If
        End If
        Conn.Close()
        TextBox8.Focus()
    End Sub
    Sub dgv3()
        konek.Koneksi()
        da = New OleDbDataAdapter("Select * from sandi", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "sandi")
        DataGridView3.DataSource = (ds.Tables("sandi"))
        DataGridView3.Enabled = True
        Conn.Close()
    End Sub
    Private Sub isiTextBox2(ByVal x As Integer)
        Try
            TextBox8.Text = DataGridView3.Rows(x).Cells(1).Value
            TextBox9.Text = DataGridView3.Rows(x).Cells(2).Value
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        isiTextBox2(e.RowIndex)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim pesan As Integer
        Dim hapus As String
        If TextBox8.Text = "" Or TextBox9.Text = "" Then
            MsgBox("Field kosong terdeteksi", vbExclamation, "Peringatan")
        Else
            konek.Koneksi()
            cmd = New OleDbCommand("SELECT * FROM sandi WHERE nama='" & TextBox8.Text & "'", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                pesan = MsgBox("Anda yakin akan hapus sandi " + TextBox8.Text, vbExclamation + vbYesNo, "Peringatan")
                If pesan = vbNo Then
                    Exit Sub
                End If
                hapus = "DELETE FROM sandi WHERE nama='" & TextBox8.Text & "'"
                konek.jalankansql(hapus)
                MsgBox("Data telah dihapus", vbExclamation)
                dgv3()
                DataGridView3.Refresh()
                TextBox8.Focus()
                bersih3()
            Else
                MsgBox("Data Admin tidak dikenal", vbExclamation, "Peringatan")
            End If
        End If
        Conn.Close()
        TextBox8.Focus()
    End Sub
    Sub bersih3()
        TextBox8.Clear()
        TextBox9.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        DataGridView2.DataSource = DBNull.Value
        TextBox5.Focus()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Field kosong terdeteksi", vbExclamation, "Peringatan")
        Else
            konek.Koneksi()
            cmd = New OleDbCommand("SELECT nik FROM data_karyawan WHERE nik='" & TextBox5.Text & "'", Conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                dgv2()
            Else
                MsgBox("Data NIK tidak ada", vbExclamation, "Perhatian")
            End If
        End If
        Conn.Close()
    End Sub
    Sub dgv2()
        konek.Koneksi()
        da = New OleDbDataAdapter("SELECT data_karyawan.nik, data_karyawan.nama_lkp, masuk.tgl,pulang.jam_krj from data_karyawan inner join (masuk inner join pulang on masuk.id_masuk=pulang.id_masuk) on data_karyawan.nik=masuk.nik where  data_karyawan.nik = '" & TextBox5.Text & "' and masuk.tgl between #" & TextBox6.Text & "# and #" & TextBox7.Text & "# ", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "abc")
        DataGridView2.DataSource = (ds.Tables("abc"))
        DataGridView2.Enabled = True
        Conn.Close()
    End Sub
    Sub excel()
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Microsoft.Office.Interop.Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        For i = 0 To DataGridView2.RowCount - 2
            For j = 0 To DataGridView2.ColumnCount - 1
                For k As Integer = 1 To DataGridView2.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView2.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = DataGridView2(j, i).Value.ToString()
                Next
            Next
        Next
        SaveFileDialog1.Filter = "XLSX Files (*.xlsx)|*.xlsx"
        If SaveFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then xlWorkSheet.SaveAs(SaveFileDialog1.FileName)
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If DataGridView2.Rows.Count Then
            excel()
        Else
            MsgBox("Tidak ada data untuk disimpan", vbExclamation, "Perhatian")
        End If
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If e.KeyChar = Chr(13) Then Button5.Focus()
    End Sub
End Class
