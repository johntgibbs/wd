Public Class EForklift1

    Private Sub tp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tp1.Click, tp2.Click, tp3.Click,
        tp4.Click, tp5.Click, tp6.Click, tp7.Click, tp8.Click, tp9.Click, tp0.Click
        Dim s As String = ""
        s = Microsoft.VisualBasic.Right(sender.ToString, 1)
        userid.Text = userid.Text & s
    End Sub

    Private Sub xit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub

    Private Sub bs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bs.Click
        If Len(userid.Text) < 2 Then
            userid.Text = ""
        Else
            userid.Text = Mid(userid.Text, 1, Len(userid.Text) - 1)
        End If
    End Sub

    Private Sub EForklift1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        userid.Text = ""
        emess.Visible = False
        userid.Focus()
    End Sub

    Private Sub EForklift1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Wdb.Close()
        'MsgBox("bye")
        End
    End Sub

    Private Sub EForklift1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If Asc(e.KeyChar) = Keys.Escape Then Me.Close()
    End Sub

    Private Sub EForklift1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo vberror
        vberror_log = "\\bbc-01-prodtrk\wd\testlogs\mobileerrors.txt"
        userid.Text = ""
        emess.Visible = False
        'Me.bbsr.Text = "ODBC;DATABASE=WDRacks;DSN=wdracks"
        Me.bbsr.Text = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
        WDbbsr = Me.bbsr.Text
        'logdir = "v:\testlogs\"
        logdir = "\\bbc-01-prodtrk\wd\pallogs\"
        'dtdock.Text = "DOCK " & Command()
        SRFlag = True
        Wdb = CreateObject("ADODB.Connection")
        Wdb.Open(WDbbsr)

        labfmtfile = "S:\wd\bin\labfmt.txt"
        Call load_labpics()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "form_load", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: form_load: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If

    End Sub

    Private Sub EForklift1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Label1.Left = (Me.Width - Label1.Width) * 0.5
        userid.Left = (Me.Width - userid.Width) * 0.5
        Button1.Left = (Me.Width - Button1.Width) * 0.5
        emess.Left = (Me.Width - emess.Width) * 0.5
        xit.Left = Me.Width - xit.Width
        tp2.Left = (Me.Width - tp2.Width) * 0.5
        tp5.Left = tp2.Left
        tp8.Left = tp2.Left
        tp0.Left = tp2.Left
        bs.Left = (Me.Width - bs.Width) * 0.5
        tp1.Left = tp2.Left - (tp1.Width * 1.5)
        tp4.Left = tp1.Left
        tp7.Left = tp1.Left
        tp3.Left = tp2.Left + (tp3.Width * 1.5)
        tp6.Left = tp3.Left
        tp9.Left = tp3.Left
    End Sub

    Private Sub userid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles userid.TextChanged
        If Len(userid.Text) >= 6 Then
            userid.Text = Mid(userid.Text, 1, 6)
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Len(userid.Text) <> 6 Then
            emess.Visible = True
        Else
            emess.Visible = False
            WDUserId = Me.userid.Text
            Eforklift2.Show()
        End If
    End Sub

End Class
