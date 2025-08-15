Public Class tpad

    Private Sub tpa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpa.Click, tpb.Click, tpc.Click, tpd.Click, tpe.Click, tpf.Click, tpg.Click, tph.Click, tpi.Click, tpj.Click, tpk.Click, tpl.Click, tpm.Click, tpn.Click, tpo.Click, tpp.Click, tpq.Click, tpr.Click, tps.Click, tpt.Click, tpu.Click, tpv.Click, tpw.Click, tpx.Click, tpy.Click, tpz.Click
        Dim s As String = ""
        s = Microsoft.VisualBasic.Right(sender.ToString, 1)
        padvalue.Text = padvalue.Text & s
    End Sub

    Private Sub tpn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpn1.Click, tpn2.Click, tpn3.Click, tpn4.Click, tpn5.Click, tpn6.Click, tpn7.Click, tpn8.Click, tpn9.Click, tpn0.Click
        Dim s As String = ""
        s = Microsoft.VisualBasic.Right(sender.ToString, 1)
        padvalue.Text = padvalue.Text & s
    End Sub

    Private Sub bsa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bsa.Click
        If Len(padvalue.Text) < 2 Or Me.Text = "BarCode" Then
            padvalue.Text = ""
        Else
            padvalue.Text = Mid(padvalue.Text, 1, Len(padvalue.Text) - 1)
        End If
    End Sub

    Private Sub eora_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles eora.Click
        padvalue.Text = padvalue.Text + "EOR"
    End Sub

    Private Sub xit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub

    Private Sub tpad_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        'Me.Close()
    End Sub

    Private Sub padvalue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles padvalue.TextChanged        
        eora.Visible = False
        tpa.Visible = False
        tpb.Visible = False
        tpc.Visible = False
        tpd.Visible = False
        tpe.Visible = False
        tpf.Visible = False
        tpg.Visible = False
        tph.Visible = False
        tpi.Visible = False
        tpj.Visible = False
        tpk.Visible = False
        tpl.Visible = False
        tpm.Visible = False
        tpn.Visible = False
        tpo.Visible = False
        tpp.Visible = False
        tpq.Visible = False
        tpr.Visible = False
        tps.Visible = False
        tpt.Visible = False
        tpu.Visible = False
        tpv.Visible = False
        tpw.Visible = False
        tpx.Visible = False
        tpy.Visible = False
        tpz.Visible = False
        tpn0.Visible = False
        tpn1.Visible = False
        tpn2.Visible = False
        tpn3.Visible = False
        tpn4.Visible = False
        tpn5.Visible = False
        tpn6.Visible = False
        tpn7.Visible = False
        tpn8.Visible = False
        tpn9.Visible = False
        If Me.Text = "UserID" Or Me.Text = "Plate" Then
            padvalue.Text = Trim(padvalue.Text)
            If Len(padvalue.Text) >= 6 Then
                Exit Sub
            End If
            tpn0.Visible = True
            tpn1.Visible = True
            tpn2.Visible = True
            tpn3.Visible = True
            tpn4.Visible = True
            tpn5.Visible = True
            tpn6.Visible = True
            tpn7.Visible = True
            tpn8.Visible = True
            tpn9.Visible = True
        End If
        If Me.Text = "BarCode" Then
            If Len(padvalue.Text) >= 16 Then
                padvalue.Text = Mid(padvalue.Text, 1, 16)
                Exit Sub
            End If
            If Len(padvalue.Text) = 3 Or Len(padvalue.Text) = 10 Or Len(padvalue.Text) = 12 Then
                padvalue.Text = padvalue.Text & " "
            End If
            If Len(padvalue.Text) >= 12 And Len(padvalue.Text) < 14 Then
                eora.Visible = True
                tpn0.Visible = True
                tpn1.Visible = True
                tpn2.Visible = True
                tpn3.Visible = True
                tpn4.Visible = True
                tpn5.Visible = True
                tpn6.Visible = True
                tpn7.Visible = True
                tpn8.Visible = True
                tpn9.Visible = True
            End If
            If Len(padvalue.Text) > 10 And Len(padvalue.Text) < 13 Then
                tpa.Visible = True
                tpb.Visible = True
                tpc.Visible = True
                tpd.Visible = True
                tpe.Visible = True
                tpf.Visible = True
                tpg.Visible = True
                tph.Visible = True
                tpi.Visible = True
                tpj.Visible = True
                tpk.Visible = True
                tpl.Visible = True
                tpm.Visible = True
                tpn.Visible = True
                tpo.Visible = True
                tpp.Visible = True
                tpq.Visible = True
                tpr.Visible = True
                tps.Visible = True
                tpt.Visible = True
                tpu.Visible = True
                tpv.Visible = True
                tpw.Visible = True
                tpx.Visible = True
                tpy.Visible = True
                tpz.Visible = True
            Else
                tpn0.Visible = True
                tpn1.Visible = True
                tpn2.Visible = True
                tpn3.Visible = True
                tpn4.Visible = True
                tpn5.Visible = True
                tpn6.Visible = True
                tpn7.Visible = True
                tpn8.Visible = True
                tpn9.Visible = True
            End If
        End If
        If Me.Text = "Qty" Then
            padvalue.Text = Trim(padvalue.Text)
            If Len(padvalue.Text) >= 6 Then
                padvalue.Text = Val(Mid(padvalue.Text, 1, 6))
                Exit Sub
            End If
            tpn0.Visible = True
            tpn1.Visible = True
            tpn2.Visible = True
            tpn3.Visible = True
            tpn4.Visible = True
            tpn5.Visible = True
            tpn6.Visible = True
            tpn7.Visible = True
            tpn8.Visible = True
            tpn9.Visible = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim i As Integer = 0
        Dim f As String = "DockTruck6"
        f = fname.Text
        For i = 0 To Application.OpenForms.Count - 1
            If f = Application.OpenForms(i).Name Then
                Call paste_it(Application.OpenForms(i), Me.cname.Text, Me.padvalue.Text)
                Exit For
            End If
        Next
        Me.Close()
    End Sub

    Sub paste_it(ByVal fn As Form, ByVal cn As String, ByVal stext As String)
        Dim ctrl As Control
        For Each ctrl In fn.Controls
            If (ctrl.GetType Is GetType(TextBox)) Then
                Dim txt As TextBox = CType(ctrl, TextBox)
                If txt.Name = cn Then
                    ctrl.Text = stext
                End If
            End If
        Next
    End Sub

    Private Sub tpad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub trig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trig.Click

    End Sub

    Private Sub trig_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles trig.TextChanged
        padvalue.Text = trig.Text
    End Sub
End Class