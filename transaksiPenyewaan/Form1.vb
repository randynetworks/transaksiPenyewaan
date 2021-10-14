Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cb_jenis_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_jenis.SelectedIndexChanged
        If cb_jenis.SelectedItem = "Panggung" Then
            tb_harga_sewa.Text = 2000000
        ElseIf cb_jenis.SelectedItem = "Kursi" Then
            tb_harga_sewa.Text = 25000
        ElseIf cb_jenis.SelectedItem = "Sound System" Then
            tb_harga_sewa.Text = 900000
        ElseIf cb_jenis.SelectedItem = "Taman" Then
            tb_harga_sewa.Text = 350000
        End If
    End Sub

    Private Sub btn_keluar_Click(sender As Object, e As EventArgs) Handles btn_keluar.Click
        Close()
    End Sub

    Private Sub tb_jml_TextChanged(sender As Object, e As EventArgs) Handles tb_jml.TextChanged
        tb_sub_harga.Text = Val(tb_harga_sewa.Text) * Val(tb_jml.Text)
        tb_potongan.Text = 0.1 * Val(tb_sub_harga.Text)
        tb_sub_grand.Text = Val(tb_sub_harga.Text) - Val(tb_potongan.Text)
        tb_pajak.Text = 0.1 * Val(tb_sub_grand.Text)
        tb_grand.Text = Val(tb_sub_grand.Text) + Val(tb_pajak.Text)
    End Sub

    Private Sub btn_new_Click(sender As Object, e As EventArgs) Handles btn_new.Click
        tb_kode.Clear()
        dtp_tanggal.ResetText()
        tb_nama.Clear()
        tb_alamat.Clear()
        cb_jenis.Text = ""
        tb_harga_sewa.Clear()
        tb_jml.Clear()
        tb_sub_harga.Clear()
        tb_potongan.Clear()
        tb_sub_grand.Clear()
        tb_pajak.Clear()
        tb_grand.Clear()
    End Sub
End Class
