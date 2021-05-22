
Imports System.Data.OleDb
Public Class MusteriBilgileri

    Dim tablo As New DataTable
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\merta\Desktop\eczaastog-12-05-2020ee\EzcaStogu\EczaVeritabani.mdb")

    Private Sub temizle()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        tablo.Clear()

    End Sub


    Private Sub listele()
        baglanti.Open()
        Dim adtr As New OleDbDataAdapter("Select*from musteri", baglanti)
        adtr.Fill(tablo)
        DataGridView1.DataSource = tablo
        baglanti.Close()
    End Sub


    Private Sub MusteriBilgileri_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: Bu kod satırı 'EczaVeritabaniDataSet3.musteri' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
        Me.MusteriTableAdapter.Fill(Me.EczaVeritabaniDataSet3.musteri)
        listele()
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Label1.Text = DataGridView1.CurrentRow.Cells("firma_no").Value.ToString
        Label2.Text = DataGridView1.CurrentRow.Cells("firma_adi").Value.ToString
        Label3.Text = DataGridView1.CurrentRow.Cells("tel_no").Value.ToString
        Label4.Text = DataGridView1.CurrentRow.Cells("fax").Value.ToString
        Label5.Text = DataGridView1.CurrentRow.Cells("adres").Value.ToString
        LinkLabel1.Text = DataGridView1.CurrentRow.Cells("web_site").Value.ToString
        LinkLabel2.Text = DataGridView1.CurrentRow.Cells("e_mail").Value.ToString


        TextBox1.Text = DataGridView1.CurrentRow.Cells("firma_no").Value.ToString
        TextBox2.Text = DataGridView1.CurrentRow.Cells("firma_adi").Value.ToString
        TextBox3.Text = DataGridView1.CurrentRow.Cells("tel_no").Value.ToString
        TextBox4.Text = DataGridView1.CurrentRow.Cells("fax").Value.ToString
        TextBox5.Text = DataGridView1.CurrentRow.Cells("adres").Value.ToString
        TextBox6.Text = DataGridView1.CurrentRow.Cells("web_site").Value.ToString
        TextBox7.Text = DataGridView1.CurrentRow.Cells("e_mail").Value.ToString
        TextBox8.Text = DataGridView1.CurrentRow.Cells("firma_hesap_no").Value.ToString
        TextBox9.Text = DataGridView1.CurrentRow.Cells("ıban").Value.ToString

        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        LinkLabel2.Visible = True
        LinkLabel1.Visible = True


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("insert into musteri(firma_no,firma_adi,tel_no,fax,adres,web_site,e_mail,firma_hesap_no,ıban) values ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "','" + TextBox6.Text + "','" + TextBox7.Text + "','" + TextBox8.Text + "','" + TextBox9.Text + "')", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        MessageBox.Show("Kayıt Eklendi", "Kayıt", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
        temizle()
        listele()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("update musteri set firma_no='" + TextBox1.Text + "',firma_adi='" + TextBox2.Text + "',tel_no='" + TextBox3.Text + "',fax='" + TextBox4.Text + "',adres='" + TextBox5.Text + "',web_site='" + TextBox6.Text + "',e_mail='" + TextBox7.Text + "',firma_hesap_no='" + TextBox8.Text + "',ıban='" + TextBox9.Text + "' where firma_no='" + TextBox1.Text + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        temizle()
        listele()
        MessageBox.Show("Güncellendi !")
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("delete * from musteri where firma_no='" + DataGridView1.CurrentRow.Cells("firma_no").Value.ToString() + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        MessageBox.Show("Eczane Kaydı Silinmiştir", "Siliniyor", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
        temizle()
        listele()
    End Sub
End Class