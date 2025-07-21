Public Class Form5

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        Form1.Visible = True

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cx, cy, cxp, cyp, radi, a, ag, rm, spc, ppc As Object
        cx = txtcxca.Text
        cy = txtcyca.Text

        cxp = Math.Sqrt(Math.Pow(cx, 2))
        cyp = Math.Sqrt(Math.Pow(cy, 2))

        a = Math.Atan(cyp / cxp)
        radi = (a * 180) / Math.PI
        ag = 90 - radi

        If cx > 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy < 0 Then
            ppc = "S"
        ElseIf cx > 0 And cy < 0 Then
            ppc = "S"
        End If

        If cx > 0 And cy > 0 Then
            spc = "E"
        ElseIf cx < 0 And cy > 0 Then
            spc = "O"
        ElseIf cx < 0 And cy < 0 Then
            spc = "O"
        ElseIf cx > 0 And cy < 0 Then
            spc = "E"
        End If

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        txtresultadoaca.Text = ag
        txtresultadomca.Text = rm
        txtresultadoppcca.Text = ppc
        txtresultadospccar.Text = spc

        txtresultadoaca.Text = Format(ag, "0.00")
        txtresultadomca.Text = Format(rm, "0.00")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtresultadoaca.Text = ""
        txtresultadomca.Text = ""
        txtresultadoppcca.Text = ""
        txtresultadospccar.Text = ""
        txtcxca.Text = ""
        txtcyca.Text = ""
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub txtcxca_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcxca.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim cx, cy, cxp, cyp, radi, a, ag, rm, spc, ppc As Object
        cx = txtcxa.Text
        cy = txtcya.Text

        cxp = Math.Sqrt(Math.Pow(cx, 2))
        cyp = Math.Sqrt(Math.Pow(cy, 2))

        a = Math.Atan(cyp / cxp)
        radi = (a * 180) / Math.PI
        ag = 90 - radi

        If cx > 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy < 0 Then
            ppc = "S"
        ElseIf cx > 0 And cy < 0 Then
            ppc = "S"
        End If

        If cx > 0 And cy > 0 Then
            spc = "E"
        ElseIf cx < 0 And cy > 0 Then
            spc = "O"
        ElseIf cx < 0 And cy < 0 Then
            spc = "O"
        ElseIf cx > 0 And cy < 0 Then
            spc = "E"
        End If

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        txtresultadoacar.Text = ag
        txtresultadomcar.Text = rm
        txtresultadoppccar.Text = ppc
        txtresultadospcart.Text = spc

        txtresultadoacar.Text = Format(ag, "0.00")
        txtresultadomcar.Text = Format(rm, "0.00")

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        txtresultadoacar.Text = ""
        txtresultadomcar.Text = ""
        txtresultadoppccar.Text = ""
        txtresultadospcart.Text = ""
        txtcxa.Text = ""
        txtcya.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtresultadospcap.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim m, a, ap, ppc, spc As Object
        m = txtmodulo.Text
        a = txtap.Text

        If a > 0 And a < 90 Then
            ap = 90 - a
        ElseIf a > 90 And a < 180 Then
            ap = a - 90
        ElseIf a > 180 And a < 270 Then
            ap = 90 - (a - 180)
        ElseIf a > 270 And a < 360 Then
            ap = a - 270
        End If

        If a > 0 And a < 90 Then
            ppc = "N"
        ElseIf a > 90 And a < 180 Then
            ppc = "N"
        ElseIf a > 180 And a < 270 Then
            ppc = "S"
        ElseIf a > 270 And a < 360 Then
            ppc = "S"
        End If

        If a > 0 And a < 90 Then
            spc = "E"
        ElseIf a > 90 And a < 180 Then
            spc = "O"
        ElseIf a > 180 And a < 270 Then
            spc = "O"
        ElseIf a > 270 And a < 360 Then
            spc = "E"
        End If

        txtresultadoaap.Text = ap
        txtresultadomp.Text = m
        txtresultadoppcap.Text = ppc
        txtresultadospcap.Text = spc

        If a = 90 Then
            txtresultadoppcap.Text = "N"
        ElseIf a = 180 Then
            txtresultadoppcap.Text = "O"
        ElseIf a = 270 Then
            txtresultadoppcap.Text = "S"
        ElseIf a = 360 Then
            txtresultadoppcap.Text = "E"
        End If


    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtresultadoaap.Text = ""
        txtresultadomp.Text = ""
        txtresultadoppcap.Text = ""
        txtresultadospcap.Text = ""
        txtmodulo.Text = ""
        txtap.Text = ""
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim m, ui, uj, cx, cy, cxp, cyp, a, ap, radi, ppc, spc As Object
        m = txtmodulomu.Text
        ui = txtunitarioi.Text
        uj = txtunitarioj.Text
        cx = m * ui
        cy = m * uj

        cxp = Math.Sqrt(Math.Pow(cx, 2))
        cyp = Math.Sqrt(Math.Pow(cy, 2))

        a = Math.Atan(cyp / cxp)
        radi = (a * 180) / Math.PI

        If cx > 0 And cy > 0 Then
            ap = radi
        ElseIf cx < 0 And cy > 0 Then
            ap = 180 - radi
        ElseIf cx < 0 And cy < 0 Then
            ap = 180 + radi
        ElseIf cx > 0 And cy < 0 Then
            ap = 360 - radi
        End If

        If ap > 0 And ap < 90 Then
            ap = 90 - ap
        ElseIf ap > 90 And ap < 180 Then
            ap = ap - 90
        ElseIf ap > 180 And ap < 270 Then
            ap = 90 - (ap - 180)
        ElseIf ap > 270 And ap < 360 Then
            ap = ap - 270
        End If


        If cx > 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy > 0 Then
            ppc = "N"
        ElseIf cx < 0 And cy < 0 Then
            ppc = "S"
        ElseIf cx > 0 And cy < 0 Then
            ppc = "S"
        End If

        If cx > 0 And cy > 0 Then
            spc = "E"
        ElseIf cx < 0 And cy > 0 Then
            spc = "O"
        ElseIf cx < 0 And cy < 0 Then
            spc = "O"
        ElseIf cx > 0 And cy < 0 Then
            spc = "E"
        End If

        txtresultadoamu.Text = ap
        txtresultadomou.Text = m
        txtresultadoppcmu.Text = ppc
        txtresultadospcmu.Text = spc

        txtresultadoamu.Text = Format(ap, "0.00")

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        txtresultadoamu.Text = ""
        txtresultadomou.Text = ""
        txtresultadoppcmu.Text = ""
        txtresultadospcmu.Text = ""
        txtmodulomu.Text = ""
        txtunitarioi.Text = ""
        txtunitarioj.Text = ""

    End Sub
End Class