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

        Dim A1 As Double = txtA1.Text
        Dim A2 As Double = txtA2.Text
        Dim A3 As Double = txtA3.Text
        Dim h1 As Double = txtH1.Text * 1000
        Dim h2 As Double = txth2.Text * 1000
        Dim t As Double = txtT.Text
        Dim x As Double = txtX.Text

        '1. vẽ hình dưới
        Dim pNgang1 As New cSTR_Point(P1.X, P1.Y)
        Dim pNgang2 As New cSTR_Point(P1.X + A1, P1.Y)
        AddLine(pNgang1.X, pNgang1.Y, pNgang2.X, pNgang2.Y)

        Dim pNgangD1 As New cSTR_Point(P1.X, P1.Y - t)
        Dim pNgangD2 As New cSTR_Point(P1.X + 10, P1.Y - t)

        Dim lineNgangDuoi As New cSTR_Line(pNgangD1.X, pNgangD1.Y, pNgangD2.X, pNgangD2.Y)
        AddLine(lineNgangDuoi.X1, lineNgangDuoi.Y1, lineNgangDuoi.X2, lineNgangDuoi.Y2)
        'Dim lineOffsetN As cSTR_Line = Return_Offset_Line(pNgang1, pNgang2, -t)
        'AddLine(lineOffsetN.X1, lineOffsetN.Y1, lineOffsetN.X2, lineOffsetN.Y2)
        'AddLine(P1.X, P1.Y, P1.X + A1, P1.Y)
        AddLine(P1.X + A1, P1.Y, P1.X + A1, P1.Y - x)
        'AddLine(P1.X + A1, P1.Y - x, P1.X + A1 + A2, P1.Y - h1)
        AddLine(P1.X, P1.Y, P1.X, P1.Y - t)
        'OffsetLine(P1.X + A1, P1.Y - x, P1.X + A1 + A2, P1.Y - h1, -t)
        'đoạn chéo trên
        Dim pCheo1 As New cSTR_Point(P1.X + A1, P1.Y - x)
        Dim pCheo2 As New cSTR_Point(P1.X + A1 + A2, P1.Y - h1)
        AddLine(pCheo1.X, pCheo1.Y, pCheo2.X, pCheo2.Y)
        ' offset xuống khoảng t
        Dim lineOffsetC As cSTR_Line = Return_Offset_Line(pCheo1, pCheo2, -t)
        AddLine(lineOffsetC.X1, lineOffsetC.Y1, lineOffsetC.X2, lineOffsetC.Y2)

        Dim pGiao As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineNgangDuoi, lineOffsetC)
        If pGiao IsNot Nothing Then
            AddLine(lineNgangDuoi.X1, lineNgangDuoi.Y1, pGiao.X, pGiao.Y)
            AddLine(pGiao.X, pGiao.Y, lineOffsetC.X2, lineOffsetC.Y2)
        End If

        DialogResult = Windows.Forms.DialogResult.OK
    End Sub

End Class