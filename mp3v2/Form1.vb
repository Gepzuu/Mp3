Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports WMPLib
Public Class Form1
    Private WithEvents wp As New WindowsMediaPlayer




    Private Sub DirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DirectoryToolStripMenuItem.Click
        Dim y As New FolderBrowserDialog
        If y.ShowDialog = DialogResult.OK Then
            For Each xfol As String In My.Computer.FileSystem.GetFiles(y.SelectedPath, FileIO.SearchOption.SearchTopLevelOnly, "*.mp3")
                Dim item As ListViewItem = ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(xfol))
                item.SubItems.Add(wp.newMedia(xfol).durationString)
                item.SubItems.Add(xfol)

            Next
        End If
    End Sub

    Private Sub DeleteOneItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteOneItemToolStripMenuItem.Click
        If ListView1.Items.Count > 0 Then
            Try
                Dim x As Integer = ListView1.SelectedIndices(0)
                ListView1.Items.RemoveAt(ListView1.SelectedIndices(0))
                ListView1.Items(x).Selected = True
                ListView1.Items(x).EnsureVisible()
                ListView1.Select()

            Catch ex As Exception
                ListView1.Items(0).Selected = True
                ListView1.Items(0).EnsureVisible()
                ListView1.Select()
            End Try
        End If
    End Sub







    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ContextMenuStrip1.Show(btnAdd, 0, btnAdd.Height)
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        ContextMenuStrip2.Show(btnDel, 0, btnDel.Height)
    End Sub

    Private Sub btnPl_Click(sender As Object, e As EventArgs) Handles btnPl.Click
        ContextMenuStrip3.Show(btnPl, 0, btnPl.Height)
    End Sub



    Private Sub btnPlay_Click(sender As Object, e As EventArgs) Handles btnPlay.Click

        If wp.playState = WMPPlayState.wmppsPaused Then
            wp.controls.play()
            Timer1.Start()
        Else
            If ListView1.SelectedIndices.Count > 0 Then
                wp.URL = ListView1.Items(ListView1.SelectedIndices(0)).SubItems(2).Text
                wp.controls.play()
                ListView1.Select()
                Timer1.Start()
            End If

        End If
    End Sub

    Private Sub btnPause_Click(sender As Object, e As EventArgs) Handles btnPause.Click
        If wp.playState = WMPPlayState.wmppsPlaying Then
            wp.controls.pause()
            Timer1.Stop()
        End If
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        If wp.playState = WMPPlayState.wmppsPlaying Or WMPPlayState.wmppsPaused Then
            wp.controls.stop()
            Timer1.Stop()
            Label7.Text = "00:00"
            Label6.Text = "00:00"
            Guna2TrackBar1.Value = 0
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If ListView1.SelectedIndices.Count > 0 Then
            Try
                Dim g As Integer = ListView1.SelectedIndices(0)
                g += 1
                ListView1.Items(g).Selected = True
                ListView1.Items(g).EnsureVisible()
                ListView1.Select()
            Catch ex As Exception
                ListView1.Items(0).Selected = True
                ListView1.Items(0).EnsureVisible()
                ListView1.Select()
            End Try
            btnPlay.PerformClick()
        End If
    End Sub

    Private Sub Guna2TrackBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles Guna2TrackBar1.Scroll
        wp.controls.currentPosition = Guna2TrackBar1.Value

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label7.Text = wp.controls.currentPositionString
        Label6.Text = wp.controls.currentItem().durationString
        Guna2TrackBar1.Value = wp.controls.currentPosition
        Guna2TrackBar1.Maximum = wp.controls.currentItem().duration


    End Sub


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

        Label1.Text = Date.Now.ToString("hh:mm")
    End Sub

    Private Sub Label10_MouseHover(sender As Object, e As EventArgs) Handles Label10.MouseHover
        Label10.ForeColor = Color.Lavender
        Label10.BackColor = Color.Navy

    End Sub

    Private Sub Label10_MouseLeave(sender As Object, e As EventArgs) Handles Label10.MouseLeave
        Label10.ForeColor = Color.Lavender
        Label10.BackColor = Color.Black
    End Sub


    Private Sub Guna2TrackBar2_Scroll(sender As Object, e As ScrollEventArgs) Handles Guna2TrackBar2.Scroll
        wp.settings.volume = Guna2TrackBar2.Value
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wp = New WindowsMediaPlayer
        Timer2.Enabled = True
    End Sub

    Private Sub AddPlaylistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddPlaylistToolStripMenuItem.Click
        Dim s As New SaveFileDialog
        s.Filter = "Tutorial Playlist|*.tpp"
        s.Title = "Save Playlist"
        If s.ShowDialog = DialogResult.OK Then
            Dim mywriter As New IO.StreamWriter(s.FileName)
            For Each item As ListViewItem In ListView1.Items
                mywriter.WriteLine(item.Text + "," + item.SubItems(2).Text + "," + item.SubItems(2).Text)
            Next
            mywriter.Close()
        End If

    End Sub

    Private Sub PlaylistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlaylistToolStripMenuItem.Click
        Dim k As New OpenFileDialog
        k.Filter = "Tutorial Playlist|*.tpp"
        k.Title = "Load Playlist"
        If k.ShowDialog = DialogResult.OK Then
            Dim firstline() As String = IO.File.ReadAllLines(k.FileName)
            For Each km As String In firstline
                Dim data() As String = km.Split(","c)
                Dim item As ListViewItem = ListView1.Items.Add(data(0))
                item.SubItems.Add(data(1))
                item.SubItems.Add(data(2))

            Next
            btnPlay.PerformClick()
            ListView1.Select()
        End If
    End Sub

    Private Sub AddFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddFilesToolStripMenuItem.Click
        Dim x As New OpenFileDialog
        x.Filter = "Audio Files|*.mp3;*.wma;*.wav"
        x.Multiselect = True
        x.Title = "Open Files"
        If x.ShowDialog = DialogResult.OK Then
            For Each file As String In x.FileNames
                Dim item As ListViewItem = ListView1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(file))
                item.SubItems.Add(wp.newMedia(file).durationString)
                item.SubItems.Add(file)

            Next
            ListView1.Items(ListView1.Items.Count - 1).Selected = True
            ListView1.Items(ListView1.Items.Count - 1).EnsureVisible()
            ListView1.Select()
            btnPlay.PerformClick()

        End If
    End Sub

    Private Sub DeleteAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteAllToolStripMenuItem.Click
        If ListView1.SelectedIndices.Count > 0 Then
            ListView1.Items.Clear()
        End If
    End Sub

    Private Sub Guna2Shapes2_Click(sender As Object, e As EventArgs) Handles Guna2Shapes2.Click

    End Sub

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs) Handles Guna2PictureBox1.Click

    End Sub

    Private Sub Guna2Shapes1_Click(sender As Object, e As EventArgs) Handles Guna2Shapes1.Click

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Me.Close()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label1.Text = Date.Now.ToString("hh:mm")
    End Sub
End Class
