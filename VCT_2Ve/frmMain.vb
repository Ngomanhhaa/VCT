Imports System.Drawing
Imports System.Drawing.Text
Imports System.IO.Ports
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
        've truc
        Dim pTrx1 As New cSTR_Point(P1.X, P1.Y + 450)
        Dim pTrx2 As New cSTR_Point(P1.X, P1.Y - h1 - 600)
        AddLine(pTrx1.X, pTrx1.Y, pTrx2.X, pTrx2.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTrx1 As cSTR_Line = Return_Offset_Line(pTrx1, pTrx2, A1)
        AddLine(lineOffsetTrx1.X1, lineOffsetTrx1.Y1, lineOffsetTrx1.X2, lineOffsetTrx1.Y2, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTrx2 As cSTR_Line = Return_Offset_Line(pTrx1, pTrx2, A1 + A2)
        AddLine(lineOffsetTrx2.X1, lineOffsetTrx2.Y1, lineOffsetTrx2.X2, lineOffsetTrx2.Y2, SYS_LAYER_AXIS_NAME)

        Dim pTry1 As New cSTR_Point(P1.X - 300, P1.Y)
        Dim pTry2 As New cSTR_Point(P1.X + A1 + A2 + 660, P1.Y)
        AddLine(pTry1.X, pTry1.Y, pTry2.X, pTry2.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTry As cSTR_Line = Return_Offset_Line(pTry1, pTry2, -h1)
        AddLine(lineOffsetTry.X1, lineOffsetTry.Y1, lineOffsetTry.X2, lineOffsetTry.Y2, SYS_LAYER_AXIS_NAME)

        ' vẽ hình dưới
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
        'AddLine(lineOffsetC.X1, lineOffsetC.Y1, lineOffsetC.X2, lineOffsetC.Y2)

        Dim pGiao As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineNgangDuoi, lineOffsetC)
        AddLine(lineNgangDuoi.X1, lineNgangDuoi.Y1, pGiao.X, pGiao.Y)
        Dim pGiao2 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffsetC, lineOffsetTry)
        AddLine(pGiao.X, pGiao.Y, pGiao2.X, pGiao2.Y)

        'dim
        AddDimX(P1.X, P1.X + A1, P1.Y - h1 - 600 - 50, -175)
        AddDimX(P1.X + A1, P1.X + A1 + A2, P1.Y - h1 - 600 - 50, -175)

        AddDimY(P1.X + A1 + 200, P1.Y, P1.Y - x, 175)
        AddDimY(P1.X + A1 + +A2 + 600 + 200, P1.Y, P1.Y - h1, 175)
        AddDimY(P1.X - 50, P1.Y, P1.Y - t, -175)

        'Add_CosCD(P1.X - 50 - 175 + 50, P1.Y, h1)
        'Add_CosCD(X2 + 100, Y0_2 + 100, h1)

        'thep
        'Add_SteelTop(P1.X, P1.Y, P1.X + A1, P1.Y, True)
        'Add_SNode(P1.X + 150, P1.Y - 35, 25)

        AddLine(P1.X + 35, P1.Y - t + Abv, P1.X + A1 - 35, P1.Y - t + Abv, SYS_LAYER_STEEL_NAME)
        Dim Point_Array As New ArrayList()
        Point_Array.Add(New Point2d(P1.X + Abv, P1.Y - t + Abv))        ' điểm dưới trái
        Point_Array.Add(New Point2d(P1.X + Abv, P1.Y - Abv))            ' điểm trên trái
        Point_Array.Add(New Point2d(P1.X + A1 - Abv, P1.Y - Abv))       ' điểm trên phải
        Point_Array.Add(New Point2d(P1.X + A1 - Abv, P1.Y - t - 100))   ' điểm dưới phải
        Add_PLine(Point_Array, SYS_LAYER_STEEL_NAME)

        'Dim pGiao3 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang()

        Dim Point_Array1 As New ArrayList()
        Point_Array1.Add(New Point2d(P1.X + A1 - 93, P1.Y - Abv))
        Point_Array1.Add(New Point2d(P1.X + A1 - 93 - 200, P1.Y - Abv))
        Point_Array1.Add(New Point2d(P1.X + A1 - 93 - 200 + 50, P1.Y - Abv - 30))
        Add_PLine(Point_Array1, SYS_LAYER_STEEL_NAME)

        DialogResult = Windows.Forms.DialogResult.OK
    End Sub

End Class