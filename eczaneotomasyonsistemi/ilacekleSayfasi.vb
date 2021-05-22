Imports System.Data.OleDb
Public Class OzelilacEkle
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\merta\Desktop\eczaastog-06-05-2020ee\EzcaStogu\EczaVeritabani.mdb")
    Private Sub temizle()
        text1.Text = ""
        text2.Text = ""
        text3.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""

    End Sub
    Dim tablo As New DataTable
    Private Sub Listele()
        baglanti.Open()
        Dim adpr As New OleDbDataAdapter("select * from ozel_ilac ", baglanti)
        adpr.Fill(tablo)
        Ozelİlac.DataGridView1.DataSource = tablo
        baglanti.Close()
    End Sub

    Private Sub BtnEkle_Click(sender As Object, e As EventArgs) Handles BtnEkle.Click

        baglanti.Open()
        Dim komut As New OleDbCommand("insert into ozel_ilac(ilac_no, ilac_adi, ilac_turu, recete_turu, raf_omru, etkin_madde)values('" + TextBox1.Text + "','" + text1.Text + "','" + ComboBox2.Text + "','" + ComboBox1.Text + "','" + text2.Text + "','" + text3.Text + "')", baglanti)
        If text1.Text = "" Then
            MessageBox.Show("İLAÇ ADI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf TextBox1.Text = "" Then
            MessageBox.Show("İLAÇ NUMARASI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text2.Text = "" Then
            MessageBox.Show("RAF ÖMRÜ ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text3.Text = "" Then
            MessageBox.Show("ETKİN MADDE ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf ComboBox2.Text = "" Then
            MessageBox.Show("İLAÇ TÜRÜ ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf ComboBox1.Text = "" Then
            MessageBox.Show("REÇETE TÜRÜ ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        Else
            komut.ExecuteNonQuery()
            baglanti.Close()
            MessageBox.Show("Kayıt Eklendi", "Kayıt", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
            temizle()
            Listele()

        End If
    End Sub

    Private Sub BtnKapat_Click(sender As Object, e As EventArgs) Handles BtnKapat.Click
        Me.Close()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub OzelilacEkle_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class