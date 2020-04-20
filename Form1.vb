Imports System.Data.OleDb
Imports AxWMPLib
Public Class form1
    Function Data(ByVal FilePath As String, ByVal sql As String) As DataTable
        Dim Connection As New OleDb.OleDbConnection
        Dim Command As New OleDb.OleDbCommand
        Dim DA As New OleDb.OleDbDataAdapter
        Dim DT As New DataTable
        Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & FilePath & ";Extended Properties=Excel 8.0"
        Command.Connection = Connection
        Command.CommandText = sql
        DA.SelectCommand = Command
        DA.Fill(DT)
        Return DT
    End Function
    Private Sub Forml_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Shown
        Dim StartPath As String = Environment.CurrentDirectory
        Dim DT As New DataTable
        DT = Data(StartPath & "\content\database\database.xls", "select distinct الانجاز from [data$]")
        achievements.DataSource = DT
        achievements.DisplayMember = DT.Columns(0).ToString
        achievements.ValueMember = DT.Columns(0).ToString
        achievements.SelectedIndex = 0
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles achievements.SelectedIndexChanged
        Try
            Dim StartPath As String = Environment.CurrentDirectory
            Dim DT As New DataTable
            DT = Data(StartPath & "\content\database\database.xls", "select الاسم from [data$] where الانجاز='" & achievements.SelectedValue & "' ")
            projects.DataSource = DT
            projects.DisplayMember = DT.Columns(0).ToString
            projects.ValueMember = DT.Columns(0).ToString
            projects.SelectedIndex = 8
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles projects.SelectedIndexChanged
        Try
            Dim StartPath As String = Environment.CurrentDirectory
            Dim DT As New DataTable
            DT = Data(StartPath & "\content\database\database.xls", "select * from [data$] where الاسم= '" & projects.SelectedValue & "' ")
            cost.Text = DT.Rows(0).Item(2).ToString
            takentime.Text = DT.Rows(0).Item(3).ToString
            goals.Text = DT.Rows(0).Item(4).ToString
            description.Text = DT.Rows(0).Item(5).ToString
            returns.Text = DT.Rows(0).Item(6).ToString
            Dim background As String = DT.Rows(0).Item(1).ToString & ".png"
            PictureBox1.Image = Image.FromFile(StartPath & "\content\backgrounds\" & background)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles link.Click
        Dim StartPath As String = Environment.CurrentDirectory
        Dim DT As New DataTable
        DT = Data(StartPath & "\content\database\database.xls", "select * from [data$] where الاسم= '" & projects.SelectedValue & "' ")
        If achievements.SelectedIndex = 1 Then
            If projects.SelectedIndex = 0 Then
                Process.Start("https://egy-map.com/project/قناة-السويس-الجديدة")
            End If
            If projects.SelectedIndex = 1 Then
                Process.Start("https://egy-map.com/project/أنفاق-قناة-السويس")
            End If
            If projects.SelectedIndex = 2 Then
                Process.Start("https://egy-map.com/project/المتحف-المصري-الكبير")
            End If
            If projects.SelectedIndex = 3 Then
                Process.Start("https://egy-map.com/project/مدينة-الجلود-بالروبيكى")
            End If
            If projects.SelectedIndex = 4 Then
                Process.Start("https://egy-map.com/project/مشروع-المليون-ونصف-المليون-فدان")
            End If
            If projects.SelectedIndex = 5 Then
                Process.Start("https://egy-map.com/project/مدينة-دمياط-للأثاث")
            End If
            If projects.SelectedIndex = 6 Then
                Process.Start("https://egy-map.com/project/بنبان-لتوليد-الكهرباء-من-الطاقة-الشمسية")
            End If
            If projects.SelectedIndex = 7 Then
                Process.Start("https://egy-map.com/project/محطة-الضبعة-النووية")
            End If
            If projects.SelectedIndex = 8 Then
                Process.Start("https://ar.wikipedia.org/wiki/العاصمة_الإدارية_(مصر)")
                Process.Start("https://www.youtube.com/watch?v=jJLQ-kDnf2w")
            End If
            If projects.SelectedIndex = 9 Then
                Process.Start("https://egy-map.com/project/مجمع-الأسمدة-الفوسفاتية-و-المركبة")
            End If
            If projects.SelectedIndex = 10 Then
                Process.Start("https://egy-map.com/project/مدينة-العلمين-الجديدة")
            End If
            If projects.SelectedIndex = 11 Then
                Process.Start("https://egy-map.com/project/قمر-الإتصالات-طيبة-TIBA-1-%E2%80%8E")
            End If
        End If
        If achievements.SelectedIndex = 0 Then
            If projects.SelectedIndex = 0 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-100-مليون-صحة")
            End If
            If projects.SelectedIndex = 1 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-حياة-كريمة")
            End If
            If projects.SelectedIndex = 2 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-نور-حياة")
            End If
            If projects.SelectedIndex = 3 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-أطفال-بلا-مأوى")
            End If
            If projects.SelectedIndex = 4 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-تكافل-وكرامة")
            End If
            If projects.SelectedIndex = 5 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-فرصة")
            End If
            If projects.SelectedIndex = 6 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-سكن-كريم")
            End If
            If projects.SelectedIndex = 7 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-اتنين-كفاية")
                Process.Start("https://www.youtube.com/watch?v=K39cCZfjAcI")
            End If
            If projects.SelectedIndex = 8 Then
                Process.Start("https://www.itida.gov.eg/Arabic/Pages/Next-Technology-Leaders.aspx")
            End If
            If projects.SelectedIndex = 9 Then
                Process.Start("https://egy-map.com/initiative/مبادرة-مصر-بلا-غارمين")
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If achievements.SelectedIndex = 0 Then
            Process.Start("https://egy-map.com/initiatives")
        Else
            Process.Start("https://egy-map.com/projects")
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Startpath As String = Environment.CurrentDirectory
        Process.Start(Startpath & "\content\videos\preview.mp4")
    End Sub
End Class
