
Imports System.Data.OleDb
Public Class PersonelBilgileri

    Dim tablo As New DataTable
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\merta\Desktop\eczaastog-12-05-2020ee\EzcaStogu\EczaVeritabani.mdb")
    Private Sub listele()
        baglanti.Open()
        Dim adtr As New OleDbDataAdapter("Select * from personel", baglanti)
        adtr.Fill(tablo)
        DataGridView1.DataSource = tablo
        baglanti.Close()

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        'Güncelleme Event'i için textboxlardan veri alınıyor 
        Txt1.Text = DataGridView1.CurrentRow.Cells("pers_adi").Value.ToString
        Txt2.Text = DataGridView1.CurrentRow.Cells("pers_soyadi").Value.ToString
        Txt3.Text = DataGridView1.CurrentRow.Cells("tc_no").Value.ToString
        txt4.Text = DataGridView1.CurrentRow.Cells("departman").Value.ToString
        ComboBox1.Text = DataGridView1.CurrentRow.Cells("unvan").Value.ToString
        ComboBox2.Text = DataGridView1.CurrentRow.Cells("cinsiyet").Value.ToString
        Txt5.Text = DataGridView1.CurrentRow.Cells("e_mail").Value.ToString
        Txt6.Text = DataGridView1.CurrentRow.Cells("dog_yeri").Value.ToString
        Txt7.Text = DataGridView1.CurrentRow.Cells("tel_no").Value.ToString

    End Sub
    Private Sub PersonelBilgileri_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: DataGridView'in Stun ve satırları renklendirildi - (Arayüz Düzenlemesi) 
        ' Me.PersonelTableAdapter.Fill(Me.EczaVeritabanıDataSet.personel)
        DataGridView1.BorderStyle = BorderStyle.None
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249)
        DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        DataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise
        DataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke
        DataGridView1.BackgroundColor = Color.White
        DataGridView1.EnableHeadersVisualStyles = False
        DataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(20, 25, 72)
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White

        listele()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        baglanti.Open()
        Dim komut As New OleDbCommand("update personel set pers_adi='" + Txt1.Text + "',pers_soyadi='" + Txt2.Text + "',e_mail='" + Txt5.Text + "',departman='" + txt4.Text + "',unvan='" + ComboBox1.Text + "',cinsiyet='" + ComboBox2.Text + "',dog_tar='" + DateTimePicker1.Text + "',dog_yeri='" + Txt6.Text + "',tel_no='" + Txt7.Text + "' where tc_no='" + Txt3.Text + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        listele()
        MessageBox.Show("Güncellendi !")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("delete *from personel where tc_no='" + DataGridView1.CurrentRow.Cells("tel_no").Value.ToString() + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        listele()
        MessageBox.Show("KAYIT SİLİNDİ!")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class