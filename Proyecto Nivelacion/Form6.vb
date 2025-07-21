Public Class Form6

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        Form1.Visible = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim m, a, cx, cy, radi, ui, uj As Double
        m = txtmoduloap.Text
        a = txtap.Text
        radi = (Math.PI * a) / 180
        cx = Math.Cos(radi) * m
        cy = Math.Sin(radi) * m

        If txtap.Text = 90 Or txtap.Text = -90 Then
            cx = 0

        ElseIf txtap.Text = 180 Or txtap.Text = -180 Then
            cy = 0

        ElseIf txtap.Text = 270 Or txtap.Text = -270 Then
            cx = 0

        ElseIf txtap.Text = 360 Or txtap.Text = -360 Then
            cy = 0
        End If

        ui = cx / m
        uj = cy / m

        txtrespuestamoduloap.Text = m
        txtrespuestaui.Text = ui
        txtrespuestauj.Text = uj

        txtrespuestaui.Text = Format(ui, "0.00")
        txtrespuestauj.Text = Format(uj, "0.00")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        txtrespuestamoduloap.Text = ""
        txtrespuestaui.Text = ""
        txtrespuestauj.Text = ""
        txtmoduloap.Text = ""
        txtap.Text = ""

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtanuglog.Text = ""
        txtmodulog.Text = ""
        txtprimerpc.Text = ""
        txtsegundopc.Text = ""
        txtrespuestamg.Text = ""
        txtrespuestauig.Text = ""
        txtrespuestaujg.Text = ""
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ag, mg, cxg, cyg, ap, radi, ppc, sppc, ui, uj As Object
        ag = txtanuglog.Text
        mg = txtmodulog.Text
        ppc = txtprimerpc.Text
        sppc = txtsegundopc.Text

        If ppc = "N" And sppc = "E" And ag = "" Then
            ag = 45
        ElseIf ppc = "E" And sppc = "N" And ag = "" Then
            ag = 45
        ElseIf ppc = "N" And sppc = "O" And ag = "" Then
            ag = 45
        ElseIf ppc = "O" And sppc = "N" And ag = "" Then
            ag = 45
        ElseIf ppc = "S" And sppc = "O" And ag = "" Then
            ag = 45
        ElseIf ppc = "O" And sppc = "S" And ag = "" Then
            ag = 45
        ElseIf ppc = "S" And sppc = "E" And ag = "" Then
            ag = 45
        ElseIf ppc = "E" And sppc = "S" And ag = "" Then
            ag = 45
        End If

        If ppc = "N" And sppc = "E" Then
            ap = 90 - ag
        ElseIf ppc = "E" And sppc = "N" Then
            ap = ag
        ElseIf ppc = "N" And sppc = "O" Then
            ap = 90 + ag
        ElseIf ppc = "O" And sppc = "N" Then
            ap = 180 - ag
        ElseIf ppc = "S" And sppc = "O" Then
            ap = 270 - ag
        ElseIf ppc = "O" And sppc = "S" Then
            ap = 180 + ag
        ElseIf ppc = "S" And sppc = "E" Then
            ap = 270 + ag
        ElseIf ppc = "E" And sppc = "S" Then
            ap = 360 - ag
        ElseIf ppc = "" And sppc = "E" And ag = "" Then
            ap = 0
        ElseIf ppc = "" And sppc = "N" And ag = "" Then
            ap = 90
        ElseIf ppc = "" And sppc = "O" And ag = "" Then
            ap = 180
        ElseIf ppc = "" And sppc = "S" And ag = "" Then
            ap = 270
        ElseIf ppc = "E" And sppc = "" And ag = "" Then
            ap = 0
        ElseIf ppc = "N" And sppc = "" And ag = "" Then
            ap = 90
        ElseIf ppc = "O" And sppc = "" And ag = "" Then
            ap = 180
        ElseIf ppc = "S" And sppc = "" And ag = "" Then
            ap = 270
        End If

        radi = (Math.PI * ap) / 180

        cxg = Math.Cos(radi) * mg
        cyg = Math.Sin(radi) * mg

        ui = cxg / mg
        uj = cyg / mg

        txtrespuestamg.Text = mg
        txtrespuestauig.Text = ui
        txtrespuestaujg.Text = uj

        txtrespuestauig.Text = Format(ui, "0.00")
        txtrespuestaujg.Text = Format(uj, "0.00")

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim cx, cy, ci, cj, ui, uj, rm As Object
        cx = txtcompxmu.Text
        cy = txtcompymu.Text
        ci = cx
        cj = cy

        If txtcompymu.Text = "" Then
            cj.Text = 0
        End If

        If txtcompxmu.Text = "" Then
            ci.Text = 0
        End If

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        ui = ci / rm
        uj = cj / rm

        txtrespuestamca.Text = rm
        txtrespuestauica.Text = ui
        txtrespuestaujca.Text = uj

        txtrespuestamca.Text = Format(rm, "0.00")
        txtrespuestauica.Text = Format(ui, "0.00")
        txtrespuestaujca.Text = Format(uj, "0.00")

    End Sub

    Private Sub Button565_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button565.Click
        txtrespuestamca.Text = ""
        txtrespuestauica.Text = ""
        txtrespuestaujca.Text = ""
        txtcompxmu.Text = ""
        txtcompymu.Text = ""

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim cx, cy, ci, cj, rm, ui, uj As Object
        cx = txtcompxv.Text
        cy = txtcompyv.Text
        ci = cx
        cj = cy
        
        If txtcompyv.Text = "" Then
            cj.Text = 0
        End If

        If txtcompxv.Text = "" Then
            ci.Text = 0
        End If

        rm = Math.Sqrt(Math.Pow(cx, 2) + Math.Pow(cy, 2))

        ui = ci / rm
        uj = cj / rm

        txtrespuestamv.Text = rm
        txtrespuestauiv.Text = ui
        txtrespuestaujv.Text = uj

        txtrespuestamv.Text = Format(rm, "0.00")
        txtrespuestauiv.Text = Format(ui, "0.00")
        txtrespuestaujv.Text = Format(uj, "0.00")

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        txtrespuestamv.Text = ""
        txtrespuestauiv.Text = ""
        txtrespuestaujv.Text = ""
        txtcompxv.Text = ""
        txtcompyv.Text = ""

    End Sub
End Class