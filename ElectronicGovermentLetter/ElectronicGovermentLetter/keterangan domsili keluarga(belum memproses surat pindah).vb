Imports word = Microsoft.Office.Interop.Word
Public Class keterangan_domsili_keluarga_belum_memproses_surat_pindah_
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        ComboBox7.Text = ""
        ComboBox8.Text = ""
        ComboBox10.Text = ""
        ComboBox9.Text = ""
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\keterangan domsili keluarga(belum memproses surat pindah).docx")
        ObjDocWord.Bookmarks("no_surat").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("tl").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("tgll").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("jk").Select()
        ObjAppWord.Selection.TypeText(ComboBox1.Text)
        ObjDocWord.Bookmarks("sp").Select()
        ObjAppWord.Selection.TypeText(ComboBox2.Text)
        ObjDocWord.Bookmarks("nama1").Select()
        ObjAppWord.Selection.TypeText(TextBox4.Text)
        ObjDocWord.Bookmarks("tl1").Select()
        ObjAppWord.Selection.TypeText(TextBox5.Text)
        ObjDocWord.Bookmarks("tgll1").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker3.Text)
        ObjDocWord.Bookmarks("jk1").Select()
        ObjAppWord.Selection.TypeText(ComboBox4.Text)
        ObjDocWord.Bookmarks("sp1").Select()
        ObjAppWord.Selection.TypeText(ComboBox3.Text)
        ObjDocWord.Bookmarks("nama2").Select()
        ObjAppWord.Selection.TypeText(TextBox6.Text)
        ObjDocWord.Bookmarks("tl2").Select()
        ObjAppWord.Selection.TypeText(TextBox7.Text)
        ObjDocWord.Bookmarks("tgll2").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker4.Text)
        ObjDocWord.Bookmarks("jk2").Select()
        ObjAppWord.Selection.TypeText(ComboBox6.Text)
        ObjDocWord.Bookmarks("sp2").Select()
        ObjAppWord.Selection.TypeText(ComboBox5.Text)
        ObjDocWord.Bookmarks("nama3").Select()
        ObjAppWord.Selection.TypeText(TextBox8.Text)
        ObjDocWord.Bookmarks("tl3").Select()
        ObjAppWord.Selection.TypeText(TextBox9.Text)
        ObjDocWord.Bookmarks("tgll3").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker5.Text)
        ObjDocWord.Bookmarks("jk3").Select()
        ObjAppWord.Selection.TypeText(ComboBox8.Text)
        ObjDocWord.Bookmarks("sp3").Select()
        ObjAppWord.Selection.TypeText(ComboBox7.Text)
        ObjDocWord.Bookmarks("nama4").Select()
        ObjAppWord.Selection.TypeText(TextBox10.Text)
        ObjDocWord.Bookmarks("tl4").Select()
        ObjAppWord.Selection.TypeText(TextBox11.Text)
        ObjDocWord.Bookmarks("tgll4").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("jk4").Select()
        ObjAppWord.Selection.TypeText(ComboBox10.Text)
        ObjDocWord.Bookmarks("sp4").Select()
        ObjAppWord.Selection.TypeText(ComboBox9.Text)
        ObjDocWord.Bookmarks("alamat1").Select()
        ObjAppWord.Selection.TypeText(TextBox12.Text)
        ObjDocWord.Bookmarks("alamat2").Select()
        ObjAppWord.Selection.TypeText(TextBox13.Text)
        ObjDocWord.Bookmarks("tgls").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)

        namafile = "D:\APP\NewDocument\keterangan domisili keluarga\" & TextBox2.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub
    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\keterangan domisili domisili keluarga\" & TextBox2.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        open()
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        reset()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        word()
    End Sub

    Private Sub keterangan_domsili_keluarga_belum_memproses_surat_pindah__Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("L")
        ComboBox1.Items.Add("P")
        ComboBox2.Items.Add("Kawin")
        ComboBox2.Items.Add("Belum Kawin")

        ComboBox4.Items.Add("L")
        ComboBox4.Items.Add("P")
        ComboBox3.Items.Add("Kawin")
        ComboBox3.Items.Add("Belum Kawin")

        ComboBox6.Items.Add("L")
        ComboBox6.Items.Add("P")
        ComboBox5.Items.Add("Kawin")
        ComboBox5.Items.Add("Belum Kawin")

        ComboBox8.Items.Add("L")
        ComboBox8.Items.Add("P")
        ComboBox7.Items.Add("Kawin")
        ComboBox7.Items.Add("Belum Kawin")

        ComboBox10.Items.Add("L")
        ComboBox10.Items.Add("P")
        ComboBox9.Items.Add("Kawin")
        ComboBox9.Items.Add("Belum Kawin")
    End Sub
End Class