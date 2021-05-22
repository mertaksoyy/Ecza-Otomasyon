Imports System.Data.OleDb
Public Class ToplamStok
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\merta\Desktop\eczaastog-12-05-2020ee\EzcaStogu\EczaVeritabani.mdb")
    Private Sub BtnEkle_Click(sender As Object, e As EventArgs) Handles BtnEkle.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("insert into stok_urun (urun_kodu,urun_adi,urun_turu,uretim_tar,son_t_tar,raf_omru,urun_adet,max_stok_kapasite,stok_durum) values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + DateTimePicker1.Text + "','" + DateTimePicker2.Text + "','" + Text4.Text + "','" + Text5.Text + "','" + Text6.Text + "','" + ComboBox1.Text + "')", baglanti)
        If Text1.Text = "" Then
            MessageBox.Show("ÜRÜN KODU BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf Text2.Text = "" Then
            MessageBox.Show("ÜRÜN ADI BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf Text3.Text = "" Then
            MessageBox.Show("ÜRÜN TÜRÜ BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf Text4.Text = "" Then
            MessageBox.Show("RAF ÖMRÜ BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf Text5.Text = "" Then
            MessageBox.Show("ÜRÜN ADETİ BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf Text6.Text = "" Then
            MessageBox.Show("MAX STOK BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        ElseIf ComboBox1.Text = "" Then
            MessageBox.Show("STOK DURUMU BOŞ BIRAKILAMAZ!")
            baglanti.Close()
        Else
            komut.ExecuteNonQuery()
            baglanti.Close()
            MessageBox.Show("Kayıt Edildi")
            tablo.Clear()
            Listele()
        End If
    End Sub

    Dim tablo As New DataTable
    Private Sub Listele()
        baglanti.Open()
        Dim adpr As New OleDbDataAdapter("select urun_kodu,urun_adi,urun_turu,uretim_tar,son_t_tar,raf_omru,urun_adet,max_stok_kapasite,stok_durum from stok_urun", baglanti)
        adpr.Fill(tablo)
        DataGridView1.DataSource = tablo
        baglanti.Close()
    End Sub

    Private Sub BtnSil_Click(sender As Object, e As EventArgs) Handles BtnSil.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("delete *from stok_urun where urun_kodu='" + DataGridView1.CurrentRow.Cells("urun_kodu").Value.ToString() + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        Listele()
    End Sub

    Private Sub ToplamStok_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.BorderStyle = BorderStyle.None
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249)
        DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        DataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise
        DataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke
        DataGridView1.BackgroundColor = Color.White

        DataGridView1.EnableHeadersVisualStyles = False
        DataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(80, 125, 79)
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White


        'TODO: Bu kod satırı 'EczaVeritabanıDataSet.stok_urun' tablosuna veri yükler. Bunu gerektiği şekilde taşıyabilir, veya kaldırabilirsiniz.
        'Me.stok_urunTableAdapter.Fill(Me.EczaVeritabanıDataSet.stok_urun)'
        Listele()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        baglanti.Open()
        Dim komut As New OleDbCommand("update stok_urun set urun_kodu='" + Text1.Text + "',urun_adi='" + Text2.Text + "',urun_turu='" + Text3.Text + "',uretim_tar='" + DateTimePicker1.Text + "',son_t_tar='" + DateTimePicker2.Text + "',raf_omru='" + Text4.Text + "',urun_adet='" + Text5.Text + "',max_stok_kapasite='" + Text6.Text + "',stok_durum='" + ComboBox1.Text + "' where urun_kodu='" + Text1.Text + "'", baglanti)
        komut.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        Listele()

        MessageBox.Show("Güncellendi !")

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        Text1.Text = DataGridView1.CurrentRow.Cells("urun_kodu").Value.ToString
        Text2.Text = DataGridView1.CurrentRow.Cells("urun_adi").Value.ToString
        Text3.Text = DataGridView1.CurrentRow.Cells("urun_turu").Value.ToString
        DateTimePicker1.Text = DataGridView1.CurrentRow.Cells("uretim_tar").Value.ToString
        DateTimePicker2.Text = DataGridView1.CurrentRow.Cells("son_t_tar").Value.ToString
        Text4.Text = DataGridView1.CurrentRow.Cells("raf_omru").Value.ToString
        Text5.Text = DataGridView1.CurrentRow.Cells("urun_adet").Value.ToString
        Text6.Text = DataGridView1.CurrentRow.Cells("max_stok_kapasite").Value.ToString
        ComboBox1.Text = DataGridView1.CurrentRow.Cells("stok_durum").Value.ToString


    End Sub
End Class