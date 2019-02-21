Imports word = Microsoft.Office.Interop.Word
Public Class skbedaidentitas
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
    End Sub
    Sub word()
        Dim ObjAppWord As New word.Application
        Dim ObjDocWord As New word.Document
        Dim namafile As String

        ObjDocWord = ObjAppWord.Documents.Open("D:\APP\Document\keterangan beda identitas.docx")
        ObjDocWord.Bookmarks("no_surat").Select()
        ObjAppWord.Selection.TypeText(TextBox1.Text)
        ObjDocWord.Bookmarks("x").Select()
        ObjAppWord.Selection.TypeText(TextBox2.Text)
        ObjDocWord.Bookmarks("pada").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("data").Select()
        ObjAppWord.Selection.TypeText(TextBox4.Text)
        ObjDocWord.Bookmarks("nama").Select()
        ObjAppWord.Selection.TypeText(TextBox5.Text)
        ObjDocWord.Bookmarks("tl").Select()
        ObjAppWord.Selection.TypeText(TextBox6.Text)
        ObjDocWord.Bookmarks("tgll").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker1.Text)
        ObjDocWord.Bookmarks("ket").Select()
        ObjAppWord.Selection.TypeText(TextBox7.Text)
        ObjDocWord.Bookmarks("datapd").Select()
        ObjAppWord.Selection.TypeText(TextBox3.Text)
        ObjDocWord.Bookmarks("nama2").Select()
        ObjAppWord.Selection.TypeText(TextBox8.Text)
        ObjDocWord.Bookmarks("tl2").Select()
        ObjAppWord.Selection.TypeText(TextBox9.Text)
        ObjDocWord.Bookmarks("tgl").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker3.Text)
        ObjDocWord.Bookmarks("datapada").Select()
        ObjDocWord.Bookmarks("tgls").Select()
        ObjAppWord.Selection.TypeText(DateTimePicker2.Text)

        namafile = "D:\APP\NewDocument\Surat Beda Identitas\" & TextBox5.Text & "-" & DateTimePicker2.Text & ".docx"
        ObjDocWord.SaveAs(namafile)
        ObjDocWord.Close()
        ObjAppWord.Quit()
    End Sub
    Sub open()
        Dim namafile As String
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim doc As New Microsoft.Office.Interop.Word.Document
        namafile = "D:\APP\NewDocument\Surat Beda Identitas\" & TextBox5.Text & "-" & DateTimePicker2.Text & ".docx"
        doc = word.Documents.Open(namafile)
        doc.Activate()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        word()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        reset()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        open()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Label8.Text = TextBox2.Text
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Label14.Text = TextBox7.Text
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub
End Class