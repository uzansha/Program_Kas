Imports System.Data.Odbc
Public Class Form1
    Dim saldoSekarang As String

    Sub TampilGrid()
        bukakoneksi()

        DA = New OdbcDataAdapter("select * From tbl_Kas", CONN)
        DS = New DataSet
        DA.Fill(DS, "tbl_Kas")
        DataGridView1.DataSource = DS.Tables("tbl_Kas")

        tutupkoneksi()
    End Sub

    Sub getSaldoSekarang()
        bukakoneksi()
        CMD = New OdbcCommand("Select * from tbl_Kas order by no desc limit 1", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            Label6.Text = "0"
        Else
            Label6.Text = RD.Item("saldo_sekarang")
            saldoSekarang = RD.Item("saldo_sekarang")
        End If

        tutupkoneksi()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TampilGrid()
        MunculCombo()
        getSaldoSekarang()
    End Sub

    Sub MunculCombo()
        ComboBox1.Items.Add("Masuk")
        ComboBox1.Items.Add("Keluar")
    End Sub
    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            If ComboBox1.Text = "Masuk" Then
                Dim saldoBaru As Integer
                Dim saldoMasuk As Integer = CInt(TextBox3.Text)
                Dim saldoTerakhir As Integer = CInt(saldoSekarang)
                saldoBaru = saldoMasuk + saldoTerakhir

                bukakoneksi()
                Dim simpan As String = " insert into tbl_Kas (no, tanggal, jenis, jumlah,keterangan,saldo_sekarang) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & saldoBaru & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input Data Berhasil")
                TampilGrid()
                KosongkanData()
                getSaldoSekarang()
                tutupkoneksi()

            ElseIf ComboBox1.Text = "Keluar" Then
                Dim saldoBaru As Integer
                Dim saldoKeluar As Integer = CInt(TextBox3.Text)
                Dim saldoTerakhir As Integer = CInt(saldoSekarang)
                saldoBaru = saldoTerakhir - saldoKeluar

                bukakoneksi()
                Dim simpan As String = " insert into tbl_Kas (no, tanggal, jenis, jumlah,keterangan,saldo_sekarang) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & saldoBaru & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input Data Berhasil")
                TampilGrid()
                KosongkanData()
                getSaldoSekarang()
                tutupkoneksi()

            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("Select * from tbl_Kas where No ='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("No Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("Tanggal")
                ComboBox1.Text = RD.Item("Jenis")
                TextBox3.Text = RD.Item("Jumlah")
                TextBox4.Text = RD.Item("Keterangan")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bukakoneksi()
        Dim edit As String = "update tbl_Kas set
        Tanggal='" & TextBox2.Text & "',
        Jenis='" & ComboBox1.Text & "',
        Jumlah='" & TextBox3.Text & "',
        Keterangan='" & TextBox4.Text & "'
        where No ='" & TextBox1.Text & "'"

        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
        TampilGrid()
        KosongkanData()
        tutupkoneksi()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yag akan dihapus dengan Masukan NIM dan Enter")
        Else
            If MessageBox.Show("Yakin akan dihapus..? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) Then
                bukakoneksi()
                Dim hapus As String = "delete from tbl_Kas where No='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                TampilGrid()
                KosongkanData()
                tutupkoneksi()
            End If
        End If
    End Sub
End Class
