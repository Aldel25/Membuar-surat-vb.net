Public Class Dashboard

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Environment.Exit(10)
    End Sub
    Private Sub DomisiliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomisiliToolStripMenuItem.Click
        If domisili_rt_.Enabled = True Then
            domisili_rt_.Show()
            Me.Hide()
        End If
    End Sub
    Private Sub PengantarIPPTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PengantarIPPTToolStripMenuItem.Click
        If pengantarippt.Enabled = True Then
            pengantarippt.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub PermohonanKomponesasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PermohonanKomponesasiToolStripMenuItem.Click
        If permohonan_kompensasi.Enabled = True Then
            permohonan_kompensasi.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratKuasaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratKuasaToolStripMenuItem.Click
        If suratkuasa.Enabled = True Then
            suratkuasa.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratKeteranganAnakKandungToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratKeteranganAnakKandungToolStripMenuItem.Click
        If sk_anakkandung.Enabled = True Then
            sk_anakkandung.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BedaIdentitasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BedaIdentitasToolStripMenuItem.Click
        If skbedaidentitas.Enabled = True Then
            skbedaidentitas.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub DomisiliAsliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomisiliAsliToolStripMenuItem.Click
        If domisili_asli_.Enabled = True Then
            domisili_asli_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub MengambilBPKBMotorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MengambilBPKBMotorToolStripMenuItem.Click
        If SKKM.Enabled = True Then
            SKKM.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub KSMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KSMToolStripMenuItem.Click
        If domsili_ksm.Enabled = True Then
            domsili_ksm.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub SuratPermohonanPendistribusianRastaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratPermohonanPendistribusianRastaToolStripMenuItem.Click
        If s_pendistribusian.Enabled = True Then
            s_pendistribusian.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratBalikNamaTanahToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratBalikNamaTanahToolStripMenuItem.Click
        If sk_baliknamatanah.Enabled = True Then
            sk_baliknamatanah.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub CeraiMatiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeraiMatiToolStripMenuItem.Click
        If sk_ceraimati.Enabled = True Then
            sk_ceraimati.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub KeluargaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluargaToolStripMenuItem.Click
        If keterangan_domsili_keluarga_belum_memproses_surat_pindah_.Enabled = True Then
            keterangan_domsili_keluarga_belum_memproses_surat_pindah_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SuratPindahSedangDiprosesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratPindahSedangDiprosesToolStripMenuItem.Click
        If keterangan_domisili_proses_.Enabled = True Then
            keterangan_domisili_proses_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub SudahTidakBertempatTinggalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SudahTidakBertempatTinggalToolStripMenuItem.Click
        If domisili_sudah_tidak_berdomisili_.Enabled = True Then
            domisili_sudah_tidak_berdomisili_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub OrganisasiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrganisasiToolStripMenuItem.Click
        If domisiliorganisasi.Enabled = True Then
            domisiliorganisasi.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub BelumMengurusSuratPindahToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BelumMengurusSuratPindahToolStripMenuItem.Click
        If domisili_belum_.Enabled = True Then
            domisili_belum_.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub JandaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JandaToolStripMenuItem.Click
        If skjanda.Enabled = True Then
            skjanda.Show()
            Me.Hide()
        End If
    End Sub
End Class