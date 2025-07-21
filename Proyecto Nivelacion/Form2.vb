Public Class Form2

    Private Sub btncalcular_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncalcular.Click
        Dim m, a, cx, cy, radi As Double
        m = txtmodulo.Text
        a = txtangulop.Text
        radi = (Math.PI * a) / 180
        cx = Math.Cos(radi) * m
        cy = Math.Sin(radi) * m
        txtcompx.Text = cx
        txtcompy.Text = cy

        If txtangulop.Text = 90 Or txtangulop.Text = -90 Then
            txtcompx.Text = 0

        ElseIf txtangulop.Text = 180 Or txtangulop.Text = -180 Then
            txtcompy.Text = 0

        ElseIf txtangulop.Text = 270 Or txtangulop.Text = -270 Then
            txtcompx.Text = 0

        ElseIf txtangulop.Text = 360 Or txtangulop.Text = -360 Then
            txtcompy.Text = 0

        End If

        txtcompx.Text = Format(cx, "0.00")
        txtcompy.Text = Format(cy, "0.00")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        Form1.Visible = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtangulop.Text = ""
        txtcompx.Text = ""
        txtcompy.Text = ""
        txtmodulo.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ag, mg, cxg, cyg, ap, radi, ppc, sppc As Object
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

        txtresultadoXg.Text = cxg
        txtresultadoYg.Text = cyg

        txtresultadoXg.Text = Format(cxg, "0.00")
        txtresultadoYg.Text = Format(cyg, "0.00")
    End Sub

    Private Sub cmdvolver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdvolver.Click
        txtanuglog.Text = ""
        txtmodulog.Text = ""
        txtprimerpc.Text = ""
        txtsegundopc.Text = ""
        txtresultadoXg.Text = ""
        txtresultadoYg.Text = ""
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim cx, cy, ci, cj As Object
        cx = txtcompxv.Text
        cy = txtcompyv.Text
        ci = cx
        cj = cy
        txtresultadocompi.Text = ci
        txtresultadocompj.Text = cj

        If txtcompyv.Text = "" Then
            txtresultadocompj.Text = 0
        End If

        If txtcompxv.Text = "" Then
            txtresultadocompi.Text = 0
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        txtcompxv.Text = ""
        txtcompyv.Text = ""
        txtresultadocompi.Text = ""
        txtresultadocompj.Text = ""
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim m, ui, uj, cx, cy As Object
        m = txtmodulomu.Text
        ui = txtunitarioi.Text
        uj = txtunitarioj.Text
        cx = m * ui
        cy = m * uj
        txtresultadocompxmu.Text = cx
        txtresultadocompymu.Text = cy
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        txtmodulomu.Text = ""
        txtunitarioi.Text = ""
        txtunitarioj.Text = ""
        txtresultadocompxmu.Text = ""
        txtresultadocompymu.Text = ""
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class