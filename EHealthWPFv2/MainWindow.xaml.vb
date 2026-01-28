Imports System.IO
Imports System.Text
Imports System.Collections.ObjectModel

Class MainWindow
    Private dataPasien As New ObservableCollection(Of Pasien)
    Private sedangEdit As Boolean = False
    Private indexEdit As Integer = -1

    Public Class Pasien
        Public Property ID As String
        Public Property Nama As String
        Public Property TanggalLahir As Date
        Public Property JenisKelamin As String
        Public Property Alamat As String
        Public Property Telepon As String
        Public Property Diagnosa As String
        Public Property TanggalDaftar As Date

        Public ReadOnly Property TanggalLahirStr As String
            Get
                Return TanggalLahir.ToString("dd/MM/yyyy")
            End Get
        End Property

        Public ReadOnly Property TanggalDaftasrStr As String
            Get
                Return TanggalDaftar.ToString("dd/MM/yyyy")
            End Get
        End Property

        Public ReadOnly Property UmurStr As String
            Get
                Return HitungUmur().ToString()
            End Get
        End Property

        Public Sub New(id As String, nama As String, tglLahir As Date,
                       jk As String, alamat As String, telp As String,
                       diagnosa As String, tglDaftar As Date)
            Me.ID = id
            Me.Nama = nama
            Me.TanggalLahir = tglLahir
            Me.JenisKelamin = jk
            Me.Alamat = alamat
            Me.Telepon = telp
            Me.Diagnosa = diagnosa
            Me.TanggalDaftar = tglDaftar
        End Sub

        Public Function HitungUmur() As Integer
            Dim umur As Integer = Date.Now.Year - TanggalLahir.Year
            If Date.Now < TanggalLahir.AddYears(umur) Then
                umur -= 1
            End If
            Return umur
        End Function
    End Class

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dtpTanggalLahir.SelectedDate = New Date(2000, 1, 1)
        dtpTanggalDaftar.SelectedDate = Date.Now

        dgvPasien.ItemsSource = dataPasien

        MuatDataDariFile()
        UpdateJumlahPasien()

        AddHandler btnTambah.Click, AddressOf btnTambah_Click
        AddHandler btnEdit.Click, AddressOf btnEdit_Click
        AddHandler btnHapus.Click, AddressOf btnHapus_Click
        AddHandler btnCari.Click, AddressOf btnCari_Click
        AddHandler btnBersihkan.Click, AddressOf btnBersihkan_Click
        AddHandler btnCetak.Click, AddressOf btnCetak_Click
    End Sub

    Private Function ValidasiInput(ByRef pesanError As String) As Boolean
        If String.IsNullOrWhiteSpace(txtNama.Text) Then
            pesanError = "Nama pasien harus diisi!"
            txtNama.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtID.Text) Then
            pesanError = "ID Pasien harus diisi!"
            txtID.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtAlamat.Text) Then
            pesanError = "Alamat harus diisi!"
            txtAlamat.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtTelepon.Text) Then
            pesanError = "Nomor telepon harus diisi!"
            txtTelepon.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtDiagnosa.Text) Then
            pesanError = "Diagnosa harus diisi!"
            txtDiagnosa.Focus()
            Return False
        End If

        ' Validasi tanggal
        If Not dtpTanggalLahir.SelectedDate.HasValue Then
            pesanError = "Tanggal lahir harus diisi!"
            Return False
        End If

        If Not dtpTanggalDaftar.SelectedDate.HasValue Then
            pesanError = "Tanggal daftar harus diisi!"
            Return False
        End If

        For i As Integer = 0 To dataPasien.Count - 1
            If sedangEdit AndAlso i = indexEdit Then
                Continue For
            End If

            If dataPasien(i).ID = txtID.Text.Trim() Then
                pesanError = "ID Pasien sudah digunakan!"
                txtID.Focus()
                Return False
            End If
        Next

        Return True
    End Function

    Private Sub btnTambah_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim pesanError As String = ""

            If Not ValidasiInput(pesanError) Then
                MessageBox.Show(pesanError, "Validasi Error",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            If sedangEdit Then
                dataPasien(indexEdit).Nama = txtNama.Text.Trim()
                dataPasien(indexEdit).TanggalLahir = dtpTanggalLahir.SelectedDate.Value
                dataPasien(indexEdit).JenisKelamin = CType(cboJenisKelamin.SelectedItem, ComboBoxItem).Content.ToString()
                dataPasien(indexEdit).Alamat = txtAlamat.Text.Trim()
                dataPasien(indexEdit).Telepon = txtTelepon.Text.Trim()
                dataPasien(indexEdit).Diagnosa = txtDiagnosa.Text.Trim()
                dataPasien(indexEdit).TanggalDaftar = dtpTanggalDaftar.SelectedDate.Value

                dgvPasien.Items.Refresh()

                MessageBox.Show("Data berhasil diupdate!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)
            Else
                Dim pasienBaru As New Pasien(
                    txtID.Text.Trim(),
                    txtNama.Text.Trim(),
                    dtpTanggalLahir.SelectedDate.Value,
                    CType(cboJenisKelamin.SelectedItem, ComboBoxItem).Content.ToString(),
                    txtAlamat.Text.Trim(),
                    txtTelepon.Text.Trim(),
                    txtDiagnosa.Text.Trim(),
                    dtpTanggalDaftar.SelectedDate.Value
                )

                dataPasien.Add(pasienBaru)

                MessageBox.Show("Data pasien berhasil ditambahkan!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)
            End If

            SimpanDataKeFile()
            UpdateJumlahPasien()
            BersihkanForm()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As RoutedEventArgs)
        Try
            If dgvPasien.SelectedItem Is Nothing Then
                MessageBox.Show("Pilih data yang akan diedit!", "Peringatan",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim pasienDipilih As Pasien = CType(dgvPasien.SelectedItem, Pasien)
            indexEdit = dataPasien.IndexOf(pasienDipilih)

            txtID.Text = pasienDipilih.ID
            txtNama.Text = pasienDipilih.Nama
            dtpTanggalLahir.SelectedDate = pasienDipilih.TanggalLahir

            For i As Integer = 0 To cboJenisKelamin.Items.Count - 1
                Dim item As ComboBoxItem = CType(cboJenisKelamin.Items(i), ComboBoxItem)
                If item.Content.ToString() = pasienDipilih.JenisKelamin Then
                    cboJenisKelamin.SelectedIndex = i
                    Exit For
                End If
            Next

            txtAlamat.Text = pasienDipilih.Alamat
            txtTelepon.Text = pasienDipilih.Telepon
            txtDiagnosa.Text = pasienDipilih.Diagnosa
            dtpTanggalDaftar.SelectedDate = pasienDipilih.TanggalDaftar

            sedangEdit = True
            txtID.IsEnabled = False
            btnTambah.Content = "Update Data"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnHapus_Click(sender As Object, e As RoutedEventArgs)
        Try
            If dgvPasien.SelectedItem Is Nothing Then
                MessageBox.Show("Pilih data yang akan dihapus!", "Peringatan",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim hasil As MessageBoxResult = MessageBox.Show(
                "Apakah Anda yakin ingin menghapus data ini?",
                "Konfirmasi Hapus",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            If hasil = MessageBoxResult.Yes Then
                Dim pasienDipilih As Pasien = CType(dgvPasien.SelectedItem, Pasien)
                dataPasien.Remove(pasienDipilih)

                SimpanDataKeFile()
                UpdateJumlahPasien()
                BersihkanForm()

                MessageBox.Show("Data berhasil dihapus!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnCari_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim keyword As String = txtCari.Text.Trim().ToLower()

            If String.IsNullOrWhiteSpace(keyword) Then
                dgvPasien.ItemsSource = dataPasien
                UpdateJumlahPasien()
                Return
            End If

            Dim hasilCari = From p In dataPasien
                            Where p.ID.ToLower().Contains(keyword) OrElse
                                  p.Nama.ToLower().Contains(keyword) OrElse
                                  p.Diagnosa.ToLower().Contains(keyword)
                            Select p

            dgvPasien.ItemsSource = hasilCari.ToList()
            lblJumlah.Text = "Ditemukan: " & hasilCari.Count() & " data"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnBersihkan_Click(sender As Object, e As RoutedEventArgs)
        BersihkanForm()
        txtCari.Text = ""
        dgvPasien.ItemsSource = dataPasien
        UpdateJumlahPasien()
    End Sub

    Private Sub btnCetak_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim namaFile As String = "Laporan_Pasien_" &
                                    Date.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
            Dim direktori As String = AppDomain.CurrentDomain.BaseDirectory & "Laporan\"

            If Not Directory.Exists(direktori) Then
                Directory.CreateDirectory(direktori)
            End If

            Dim pathFile As String = Path.Combine(direktori, namaFile)

            Dim sb As New StringBuilder()
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("            LAPORAN DATA PASIEN E-HEALTH")
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("Tanggal Cetak: " & Date.Now.ToString("dd MMMM yyyy HH:mm:ss"))
            sb.AppendLine("Total Pasien: " & dataPasien.Count)
            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine()

            Dim nomer As Integer = 1
            For Each pasien As Pasien In dataPasien
                sb.AppendLine("Pasien #" & nomer)
                sb.AppendLine(New String("-"c, 80))
                sb.AppendLine("ID Pasien      : " & pasien.ID)
                sb.AppendLine("Nama           : " & pasien.Nama)
                sb.AppendLine("Tanggal Lahir  : " & pasien.TanggalLahir.ToString("dd/MM/yyyy"))
                sb.AppendLine("Umur           : " & pasien.HitungUmur() & " tahun")
                sb.AppendLine("Jenis Kelamin  : " & pasien.JenisKelamin)
                sb.AppendLine("Alamat         : " & pasien.Alamat)
                sb.AppendLine("Telepon        : " & pasien.Telepon)
                sb.AppendLine("Diagnosa       : " & pasien.Diagnosa)
                sb.AppendLine("Tanggal Daftar : " & pasien.TanggalDaftar.ToString("dd/MM/yyyy"))
                sb.AppendLine()
                nomer += 1
            Next

            sb.AppendLine("=" & New String("="c, 80))
            sb.AppendLine("Akhir Laporan")

            File.WriteAllText(pathFile, sb.ToString())

            MessageBox.Show("Laporan berhasil dibuat!" & vbCrLf &
                          "File: " & pathFile, "Sukses",
                          MessageBoxButton.OK, MessageBoxImage.Information)

            Process.Start("explorer.exe", direktori)

        Catch ex As Exception
            MessageBox.Show("Error membuat laporan: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub UpdateJumlahPasien()
        lblJumlah.Text = "Total Pasien: " & dataPasien.Count
    End Sub

    Private Sub BersihkanForm()
        txtID.Text = ""
        txtNama.Text = ""
        dtpTanggalLahir.SelectedDate = New Date(2000, 1, 1)
        cboJenisKelamin.SelectedIndex = 0
        txtAlamat.Text = ""
        txtTelepon.Text = ""
        txtDiagnosa.Text = ""
        dtpTanggalDaftar.SelectedDate = Date.Now

        txtID.IsEnabled = True
        sedangEdit = False
        indexEdit = -1
        btnTambah.Content = "Tambah Data"

        txtNama.Focus()
    End Sub

    Private Sub SimpanDataKeFile()
        Try
            Dim namaFile As String = "data_pasien.txt"
            Dim pathFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, namaFile)

            Dim sb As New StringBuilder()

            For Each p As Pasien In dataPasien
                sb.AppendLine(String.Join("|",
                    p.ID, p.Nama,
                    p.TanggalLahir.ToString("yyyy-MM-dd"),
                    p.JenisKelamin, p.Alamat, p.Telepon, p.Diagnosa,
                    p.TanggalDaftar.ToString("yyyy-MM-dd")
                ))
            Next

            File.WriteAllText(pathFile, sb.ToString())

        Catch ex As Exception
            MessageBox.Show("Error menyimpan data: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub MuatDataDariFile()
        Try
            Dim namaFile As String = "data_pasien.txt"
            Dim pathFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, namaFile)

            If Not File.Exists(pathFile) Then
                Return
            End If

            Dim lines() As String = File.ReadAllLines(pathFile)

            dataPasien.Clear()

            For Each line As String In lines
                If Not String.IsNullOrWhiteSpace(line) Then
                    Dim data() As String = line.Split("|"c)

                    If data.Length >= 8 Then
                        Dim pasien As New Pasien(
                            data(0),
                            data(1),
                            Date.Parse(data(2)),
                            data(3),
                            data(4),
                            data(5),
                            data(6),
                            Date.Parse(data(7))
                        )

                        dataPasien.Add(pasien)
                    End If
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error memuat data: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class