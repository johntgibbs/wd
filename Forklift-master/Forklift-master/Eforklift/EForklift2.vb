Public Class Eforklift2

    Private Sub xit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub

    Private Sub Eforklift2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        List1.Items.Clear()
        List1.Items.Add("Move Pallets")
        List1.Items.Add(" ")
        List1.Items.Add("Ship Pallets")
        List1.Items.Add(" ")
        List1.Items.Add("Roller Bed")
        List1.Items.Add(" ")
        List1.Items.Add("Pallet History")
        List1.Items.Add(" ")
        List1.SelectedIndex = 0
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim s As String
        If List1.SelectedItem = "Move Pallets" Or List1.SelectedIndex = 1 Then
            'msgBox("Move pallet")
            EForklift4.apphdr.Text = "Move Pallets"
            EForklift4.Show()
        End If
        If List1.SelectedItem = "Ship Pallets" Or List1.SelectedIndex = 3 Then
            'MsgBox("ship pallet")
            EForklift3.apphdr.Text = "Ship Pallets"
            EForklift3.Show()
        End If
        If List1.SelectedItem = "Roller Bed" Or List1.SelectedIndex = 5 Then
            'MsgBox("roller bed")
            EForklift7.apphdr.Text = "Roller Bed"
            EForklift7.Show()
        End If
        If List1.SelectedItem = "Pallet History" Or List1.SelectedIndex = 7 Then
            Text1.Visible = True
            'MsgBox("pallet info")
            s = InputBox("Scan Barcode", "Pallet History")
            If Len(s) > 0 Then
                Text1.Text = ""
                Text1.Text = pallet_history_text(s)
                If Len(Text1.Text) > 0 Then histbc = s
            End If
        End If
    End Sub

    Private Sub Text1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.Click
        Text1.Visible = False
    End Sub

    Private Sub Eforklift2_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Label1.Left = (Me.Width - Label1.Width) * 0.5
        List1.Left = (Me.Width - List1.Width) * 0.5
        Button1.Left = (Me.Width - Button1.Width) * 0.5
        xit.Left = Me.Width - xit.Width
        Text1.Width = Me.Width - xit.Width
    End Sub


End Class