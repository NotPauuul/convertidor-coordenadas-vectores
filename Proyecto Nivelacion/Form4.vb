Public Class Form4

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        Form1.Visible = True

    End Sub

    Private Sub cmdvolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim mg, ppc, ag, spc, ap As Object
        mg = txtmodulog.Text
        ppc = txtprimerpc.Text
        spc = txtsegundopc.Text
        ag = txtanuglog.Text

        If ppc = "N" And spc = "E" And ag = "" Then
            ag = 45
        ElseIf ppc = "E" And spc = "N" And ag = "" Then
            ag = 45
        ElseIf ppc = "N" And spc = "O" And ag = "" Then
            ag = 45
        ElseIf ppc = "O" And spc = "N" And ag = "" Then
            ag = 45
        ElseIf ppc = "S" And spc = "O" And ag = "" Then
            ag = 45
        ElseIf ppc = "O" And spc = "S" And ag = "" Then
            ag = 45
        ElseIf ppc = "S" And spc = "E" And ag = "" Then
            ag = 45
        ElseIf ppc = "E" And spc = "S" And ag = "" Then
            ag = 45
        End If

        If ppc = "N" And spc = "E" Then
            ap = 90 - ag
        ElseIf ppc = "E" And spc = "N" Then
            ap = ag
        ElseIf ppc = "N" And spc = "O" Then
            ap = 90 + ag
        ElseIf ppc = "O" And spc = "N" Then
            ap = 180 - ag
        ElseIf ppc = "S" And spc = "O" Then
            ap = 270 - ag
        ElseIf ppc = "O" And spc = "S" Then
            ap = 180 + ag
        ElseIf ppc = "S" And spc = "E" Then
            ap = 270 + ag
        ElseIf ppc = "E" And spc = "S" Then
            ap = 360 - ag
        ElseIf ppc = "" And spc = "E" And ag = "" Then
            ap = 0
        ElseIf ppc = "" And spc = "N" And ag = "" Then
            ap = 90
        ElseIf ppc = "" And spc = "O" And ag = "" Then
            ap = 180
        ElseIf ppc = "" And spc = "S" And ag = "" Then
            ap = 270
        ElseIf ppc = "E" And spc = "" And ag = "" Then
            ap = 0
        ElseIf ppc = "N" And spc = "" And ag = "" Then
            ap = 90
        ElseIf ppc = "O" And spc = "" And ag = "" Then
            ap = 180
        ElseIf ppc = "S" And spc = "" And ag = "" Then
            ap = 270
        End If

        txtresultadoap.Text = ap
        txtresultadomg.Text = mg



    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtresultadomg.Text = ""
        txtresultadoap.Text = ""
        txtanuglog.Text = ""
        txtsegundopc.Text = ""
        txtprimerpc.Text = ""
        txtmodulog.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim cx, cy, cx2, cy2, cxp, cyp, a, radi, ap, rm As Object
        cx = txtcx.Text
        cy = txtcy.Text

        cx2 = Math.Pow(cx, 2)
        cxp = Math.Sqrt(cx2)
        cy2 = Math.Pow(cy, 2)
        cyp = Math.Sqrt(cy2)

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

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        txtresultadomc.Text = rm
        txtresultadoapc.Text = ap

        txtresultadomc.Text = Format(rm, "0.00")
        txtresultadoapc.Text = Format(ap, "0.00")







    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtresultadomc.Text = ""
        txtresultadoapc.Text = ""
        txtcx.Text = ""
        txtcy.Text = ""

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim cx, cy, cx2, cy2, cxp, cyp, a, radi, ap, rm As Object
        cx = txtci.Text
        cy = txtcj.Text

        cx2 = Math.Pow(cx, 2)
        cxp = Math.Sqrt(cx2)
        cy2 = Math.Pow(cy, 2)
        cyp = Math.Sqrt(cy2)

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

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        txtresultadomv.Text = rm
        txtapv.Text = ap

        txtresultadomv.Text = Format(rm, "0.00")
        txtapv.Text = Format(ap, "0.00")

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtresultadomv.Text = ""
        txtapv.Text = ""
        txtci.Text = ""
        txtcj.Text = ""
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        txtmodulomu.Text = ""
        txtunitarioi.Text = ""
        txtunitarioj.Text = ""
        txtresultadommu.Text = ""
        txtresultadoapmu.Text = ""
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim m, ui, uj, cx, cy, cxp, cyp, a, ap, radi As Object
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

        txtresultadoapmu.Text = ap
        txtresultadommu.Text = m

        txtresultadoapmu.Text = Format(ap, "0.00")


    End Sub
End Class