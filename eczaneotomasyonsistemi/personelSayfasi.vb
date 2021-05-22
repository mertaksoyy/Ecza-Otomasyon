
Imports System.Data.OleDb
Public Class Personel
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\merta\Desktop\eczaastog-06-05-2020ee\EzcaStogu\EczaVeritabani.mdb")
    Private Sub temizle()
        text1.Text = ""
        text2.Text = ""
        text4.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        DateTimePicker1.Text = ""
        text6.Text = ""
        text7.Text = ""
        text5.Text = ""
    End Sub
    Private Sub BtnEkle_Click(sender As Object, e As EventArgs) Handles BtnEkle.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("insert into personel(pers_adi,pers_soyadi,tc_no,departman,unvan,cinsiyet,dog_tar,dog_yeri,tel_no,e_mail)values('" + text1.Text + "','" + text2.Text + "','" + text3.Text + "','" + text4.Text + "','" + ComboBox2.Text + "','" + ComboBox1.Text + "','" + DateTimePicker1.Text + "','" + text6.Text + "','" + text7.Text + "','" + text5.Text + "')", baglanti)

        If text1.Text = "" Then
            MessageBox.Show("PERSONEL ADI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text2.Text = "" Then
            MessageBox.Show("PERSONEL SOYADI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text3.Text = "" Then
            MessageBox.Show("TC KİMLİK ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text4.Text = "" Then
            MessageBox.Show("DEPARTMAN ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text5.Text = "" Then
            MessageBox.Show("MAİL ADRESİ ALANU BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text6.Text = "" Then
            MessageBox.Show("DOĞUM YERİ ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf text7.Text = "" Then
            MessageBox.Show("TELEFONU ALANI BOŞ BIRAKILAMAZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf ComboBox1.Text = "" Then
            MessageBox.Show("CİNSİYET SEÇİNİZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        ElseIf ComboBox2.Text = "" Then
            MessageBox.Show("ÜNVAN SEÇİNİZ!", "Important Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            baglanti.Close()
        Else
            komut.ExecuteNonQuery()
            baglanti.Close()
            MessageBox.Show("Kayıt Eklendi", "Kayıt", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
            temizle()
        End If
    End Sub

    Private Sub Personel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub
End Class