Imports System.Drawing
Imports System.Drawing.Text
Imports System.Xml
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class frmMain
    Private Sub btnThoat_Click(sender As Object, e As EventArgs) Handles btnThoat.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub btnVe_Click(sender As Object, e As EventArgs) Handles btnVe.Click
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed = doc.Editor

        Dim ptRs = ed.GetPoint(vbLf & "chọn điểm chèn:")
        If ptRs.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK Then
            Exit Sub
        End If

        Dim P1 As Point3d = ptRs.Value

        'Copy file vào bản vẽ
        CopyStyle("C:\KetcauSoft\Com\KCS_STYLE.dwg")

        Dim H1 As Double = txtH1.Text
        Dim H2 As Double = txtH2.Text
        Dim H3 As Double = txtH3.Text
        Dim a1 As Double = txtA1.Text * 1000
        Dim a2 As Double = txtA2.Text * 1000
        Dim t As Double = txtT.Text
        Dim KcTr As Double = txtKcTr.Text
        Dim KcD As Double = txtKcD.Text

        '1. vẽ hình dưới
        AddLine(P1.X, P1.Y, P1.X + H1, P1.Y)
        AddLine(P1.X + H1, P1.Y, P1.X + H1 + H2, P1.Y - a1)
        AddLine(P1.X, P1.Y, P1.X, P1.Y - t)
        AddLine(P1.X + H1 + H2, P1.Y - a1, P1.X + H1 + H2 + 800, P1.Y - a1)
        OffsetLine(P1.X, P1.Y, P1.X + H1, P1.Y, -t)
        OffsetLine(P1.X + H1, P1.Y, P1.X + H1 + H2, P1.Y - a1, -t)
        OffsetLine(P1.X + H1 + H2, P1.Y - a1, P1.X + H1 + H2 + 800, P1.Y - a1, -t)
        Add_BreakLineY(P1.X + H1 + H2 + 800, P1.Y - a1, P1.Y - a1 - t, SYS_LAYER_THIN_NAME)

        'AddSteelNode(P1.X + 150, P1.Y + 35,SYS_LAYER_STEEL_NAME)

        ''vẽ hình trên
        'AddLine(P1.X, P1.Y + 2000, P1.X + H1, P1.Y + 2000)
        'AddLine(P1.X + H1, P1.Y + 2000, P1.X + H1 + H3, P1.Y + 2000 + a2)
        'AddLine(P1.X, P1.Y + 2000, P1.X, P1.Y + 2000 - t)
        'AddLine(P1.X + H1 + H3, P1.Y + 2000 + a2, P1.X + H1 + H3 + 800, P1.Y + 2000 + a2)
        'OffsetLine(P1.X, P1.Y + 2000, P1.X + H1, P1.Y + 2000, -t)
        'OffsetLine(P1.X + H1, P1.Y + 2000, P1.X + H1 + H3, P1.Y + 2000 + a2, -t)
        'OffsetLine(P1.X + H1 + H3, P1.Y + 2000 + a2, P1.X + H1 + H3 + 800, P1.Y + 2000 + a2, -t)
        'Add_BreakLineY(P1.X + H1 + H3 + 800, P1.Y + 2000 + a2, P1.Y + 2000 + a2 - t, SYS_LAYER_THIN_NAME)


        DialogResult = Windows.Forms.DialogResult.OK
    End Sub

End Class