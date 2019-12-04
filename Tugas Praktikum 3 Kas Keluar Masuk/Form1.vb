Imports System.Data.Odbc
Public Class Form1
    Public saldoSekarang As Integer
    Sub TampilData()
        mulaikoneksi()

        adapter = New OdbcDataAdapter("select * from keluar_masuk", CONN)
        d_set = New DataSet
        adapter.Fill(d_set, "keluar_masuk")
        DataGridView1.DataSource = d_set.Tables("keluar_masuk")

        stopkoneksi()
    End Sub
    Sub aturDGV()
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 60
        DataGridView1.Columns(3).Width = 60
        DataGridView1.Columns(4).Width = 120
        DataGridView1.Columns(5).Width = 80
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilData()
        ComboJenis()
        getSaldo()
        aturDGV()
    End Sub
    Sub ComboJenis()
        ComboBoxJenis.Items.Add("Masuk")
        ComboBoxJenis.Items.Add("Keluar")
    End Sub
    Sub BersihkanForm()
        TextBoxId.Text = ""
        DateTimePicker1.Text = ""
        ComboBoxJenis.Text = ""
        TextBoxJml.Text = ""
        TextBoxKet.Text = ""
    End Sub
    Sub getSaldo()
        mulaikoneksi()

        command = New OdbcCommand("select * from keluar_masuk order by Id desc limit 1", CONN)
        reader = command.ExecuteReader
        reader.Read()
        If Not reader.HasRows Then
            labelSaldo.Text = "0"
        Else
            labelSaldo.Text = reader.Item("saldo_sekarang")
            saldoSekarang = reader.Item("saldo_sekarang")
        End If

        stopkoneksi()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBoxId.Text = "" Or DateTimePicker1.Text = "" Or ComboBoxJenis.Text = "" Or TextBoxJml.Text = "" Or TextBoxKet.Text = "" Then
            MsgBox("Ada Data Kosong")
        Else
            mulaikoneksi()
            If ComboBoxJenis.Text = "Masuk" Then
                saldoSekarang = saldoSekarang + TextBoxJml.Text
                Dim simpan As String = "insert into keluar_masuk values ('" & TextBoxId.Text & "','" & DateTimePicker1.Text & "','" & ComboBoxJenis.Text & "','" & TextBoxJml.Text & "','" & saldoSekarang & "','" & TextBoxKet.Text & "')"

                command = New OdbcCommand(simpan, CONN)
                command.ExecuteNonQuery()
                MsgBox("Transaksi Berhasil")
                labelSaldo.Text = saldoSekarang

            ElseIf ComboBoxJenis.Text = "Keluar" Then
                If saldoSekarang < TextBoxJml.Text Then
                    MsgBox("Transaksi Gagal, Saldo Tidak Cukup")
                Else
                    saldoSekarang = saldoSekarang - TextBoxJml.Text
                    labelSaldo.Text = saldoSekarang
                    Dim simpan As String = "insert into keluar_masuk values ('" & TextBoxId.Text & "','" & DateTimePicker1.Text & "','" & ComboBoxJenis.Text & "','" & TextBoxJml.Text & "','" & saldoSekarang & "','" & TextBoxKet.Text & "')"

                    command = New OdbcCommand(simpan, CONN)
                    command.ExecuteNonQuery()
                    MsgBox("Transaksi Berhasil")

                End If
            End If
            TampilData()
            BersihkanForm()
            stopkoneksi()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBoxId.Text = "" Then
            MsgBox("Pilih data yang ingin dihapus dengan memasukkan Id dan Tekan ENTER")
        Else
            If MessageBox.Show("Hapus data ? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) Then
                mulaikoneksi()

                Dim bersihkan As String = "delete from keluar_masuk where id ='" & TextBoxId.Text & "'"
                command = New OdbcCommand(bersihkan, CONN)
                command.ExecuteNonQuery()

                TampilData()
                BersihkanForm()
                stopkoneksi()
            End If
        End If
    End Sub
    Private Sub TextBoxId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxId.KeyPress
        TextBoxId.MaxLength = 20

        If e.KeyChar = Chr(13) Then
            mulaikoneksi()
            command = New OdbcCommand("select * from keluar_masuk where Id ='" & TextBoxId.Text & "'", CONN)
            reader = command.ExecuteReader
            reader.Read()
            If Not reader.HasRows Then
                MsgBox("ID tidak ada, Silahkan coba lagi!")
                TextBoxId.Focus()
            Else
                TextBoxId.Text = reader.Item("Id")
                DateTimePicker1.Text = reader.Item("tanggal")
                ComboBoxJenis.Text = reader.Item("jenis")
                TextBoxJml.Text = reader.Item("jumlah")
                TextBoxKet.Text = reader.Item("keterangan")
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        mulaikoneksi()
        If TextBoxId.Text = "" Or DateTimePicker1.Text = "" Or ComboBoxJenis.Text = "" Or TextBoxJml.Text = "" Or TextBoxKet.Text = "" Then
            MsgBox("Data tidak bisa diperbarui, Silahkan isi semua data")
        Else
            Dim sunting As String = "update keluar_masuk set
        Id='" & TextBoxId.Text & "',
        tanggal='" & DateTimePicker1.Text & "',
        jenis='" & ComboBoxJenis.Text & "',
        jumlah='" & TextBoxJml.Text & "',
        keterangan='" & TextBoxKet.Text & "'
        where id ='" & TextBoxId.Text & "'"

            command = New OdbcCommand(sunting, CONN)
            command.ExecuteNonQuery()
            MsgBox("Data Diperbarui")

            TampilData()
            BersihkanForm()
            stopkoneksi()
        End If
    End Sub
End Class