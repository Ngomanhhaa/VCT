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
        Ve_Thang2V(P1.X, P1.Y)

    End Sub
    Sub Ve_Thang2V(X0 As Decimal, Y0 As Decimal)
        VeMB(X0, Y0)
        VeMC1(X0, Y0 - 3000)
        veMC2(X0 + 4000, Y0 - 3000)
    End Sub
    Sub VeMB(X0 As Double, Y0 As Double)
        Dim P1 As New cSTR_Point(X0, Y0, 0)

        Dim L1 As Double = txtL1.Text
        Dim L2 As Double = txtL2.Text
        Dim L3 As Double = txtL3.Text
        Dim h1 As Double = txtH1.Text * 1000
        Dim h2 As Double = txtH2.Text * 1000
        'Dim a As Double = txtA.Text * 1000
        Dim t As Double = txtT.Text
        Dim x As Double = txtX.Text
#Region "MB"
#Region "vẽ trục"
        Dim pTr1 As New cSTR_Point(P1.X, P1.Y - 300)
        Dim pTr2 As New cSTR_Point(P1.X, P1.Y + L1 + L2 + 300)
        AddLine(pTr1.X, pTr1.Y, pTr2.X, pTr2.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTr1 As cSTR_Line = Return_Offset_Line(pTr1, pTr2, -L1)
        AddLine(lineOffsetTr1.X1, lineOffsetTr1.Y1, lineOffsetTr1.X2, lineOffsetTr1.Y2, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTr2 As cSTR_Line = Return_Offset_Line(pTr1, pTr2, -L1 - L3)
        AddLine(lineOffsetTr2.X1, lineOffsetTr2.Y1, lineOffsetTr2.X2, lineOffsetTr2.Y2, SYS_LAYER_AXIS_NAME)

        Dim pTr3 As New cSTR_Point(P1.X - 300, P1.Y)
        Dim pTr4 As New cSTR_Point(P1.X + L1 + L3 + 300, P1.Y)
        AddLine(pTr3.X, pTr3.Y, pTr4.X, pTr4.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTr3 As cSTR_Line = Return_Offset_Line(pTr3, pTr4, L2)
        AddLine(lineOffsetTr3.X1, lineOffsetTr3.Y1, lineOffsetTr3.X2, lineOffsetTr3.Y2, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTr4 As cSTR_Line = Return_Offset_Line(pTr3, pTr4, L2 + L1)
        AddLine(lineOffsetTr4.X1, lineOffsetTr4.Y1, lineOffsetTr4.X2, lineOffsetTr4.Y2, SYS_LAYER_AXIS_NAME)

        AddDimX(P1.X, P1.X + L1, P1.Y + L1 + L2 + 300 + 50, 175)
        AddDimX(P1.X + L1, P1.X + L1 + L3, P1.Y + L1 + L2 + 300 + 50, 175)
        AddDimX(P1.X, P1.X + L1 + L3, P1.Y + L1 + L2 + 300 + 50, 350)

        AddDimY(P1.X - 300 - 50, P1.Y, P1.Y + L2, -175)
        AddDimY(P1.X - 300 - 50, P1.Y + L2, P1.Y + L2 + L1, -175)
        AddDimY(P1.X - 300 - 50, P1.Y, P1.Y + L2 + L1, -350)
#End Region
#Region "vẽ mb"
        Dim P2 As New cSTR_Point(P1.X, P1.Y + L1 + L2)
        Dim P3 As New cSTR_Point(P1.X + L1 + L3, P1.Y + L1 + L2)
        Dim P4 As New cSTR_Point(P1.X + L1 + L3, P1.Y + L2)
        Dim P5 As New cSTR_Point(P1.X + L1, P1.Y + L2)
        Dim P6 As New cSTR_Point(P1.X + L1, P1.Y)

        AddLine(P1.X, P1.Y, P2.X, P2.Y)
        AddLine(P2.X, P2.Y, P3.X, P3.Y)
        AddLine(P4.X, P4.Y, P5.X, P5.Y)
        AddLine(P5.X, P5.Y, P6.X, P6.Y)
        AddLine(P1.X, P1.Y + L2, P5.X, P5.Y)
        AddLine(P2.X + L1, P2.Y, P5.X, P5.Y)
        Add_BreakLineY(P3.X, P3.Y, P4.Y, SYS_LAYER_THIN_NAME)
        Add_BreakLineX(P1.Y, P1.X, P6.X, SYS_LAYER_THIN_NAME)
        Add_CosCD(P1.X + 200, P1.Y + 100, 0)
        Add_CosCD(P1.X + 200, P1.Y + L2 + 100, h1 / 1000)
        Add_CosCD(P4.X - 400, P4.Y + 100, h2 / 1000)
#End Region

#Region "thep"
#Region "thep L2"
        AddLine(P1.X + Abv_PhanThan, P1.Y + L2 / 2, P1.X + L1 - Abv_PhanThan, P1.Y + L2 / 2, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P1.X + Abv_PhanThan, P1.Y + L2 / 2 + 175, P1.X + L1 - Abv_PhanThan, P1.Y + L2 / 2 + 175, True)
        AddLine(P6.X - 125, P6.Y + Abv_PhanThan, P5.X - 125, P5.Y - Abv_PhanThan, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P6.X - 300, P6.Y + Abv_PhanThan, P5.X - 300, P5.Y - Abv_PhanThan, False)
#End Region
#Region "ThepL1"
        AddLine(P1.X + Abv_PhanThan, P1.Y + L2 + L1 / 2, P1.X + L1 - Abv_PhanThan, P1.Y + L2 + L1 / 2, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P1.X + Abv_PhanThan, P1.Y + L2 + L1 / 2 + 175, P1.X + L1 - Abv_PhanThan, P1.Y + L2 + L1 / 2 + 175, True)
        AddLine(P5.X - 125, P5.Y + Abv_PhanThan, P5.X - 125, P5.Y + L1 - Abv_PhanThan, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P5.X - 300, P5.Y + Abv_PhanThan, P5.X - 300, P5.Y + L1 - Abv_PhanThan, False)
#End Region
#Region "ThepL3"
        AddLine(P3.X + L1 + Abv_PhanThan, P3.Y + L2 + L1 / 2, P3.X + L3 - Abv_PhanThan, P1.Y + L2 + L1 / 2, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P3.X + L1 + Abv_PhanThan, P3.Y + L2 + L1 / 2 + 175, P3.X + L3 - Abv_PhanThan, P1.Y + L2 + L1 / 2 + 175, True)
        AddLine(P5.X - 125, P5.Y + Abv_PhanThan, P5.X - 125, P5.Y + L1 - Abv_PhanThan, SYS_LAYER_STEEL_NAME)
        Add_SteelTop(P5.X - 300, P5.Y + Abv_PhanThan, P5.X - 300, P5.Y + L1 - Abv_PhanThan, False)
#End Region
#End Region
    End Sub
#End Region
#Region "MC1"
    Sub VeMC1(X0 As Double, Y0 As Double)

        Dim P1 As New Point3d(X0, Y0, 0)
        Dim L1 As Double = txtL1.Text
        Dim L2 As Double = txtL2.Text
        Dim L3 As Double = txtL3.Text
        Dim h1 As Double = txtH1.Text * 1000
        Dim h2 As Double = txtH2.Text * 1000
        Dim t As Double = txtT.Text
        Dim x As Double = txtX.Text
#Region "khai báo chung"
        'L1 (chieu nghi)
        Dim Fi1 As Double = cbxFi1L1.Text
        Dim Fi2 As Double = cbxFi2L1.Text
        Dim Fi3 As Double = cbxFi3L1.Text
        Dim Fi4 As Double = cbxFi4L1.Text
        Dim a1 As Double = txta1L1.Text
        Dim a2 As Double = txta2L1.Text
        Dim a3 As Double = txta3L1.Text
        Dim a4 As Double = txta4L1.Text

        'L2
        Dim Fi5 As Double = cbxFi1L2.Text
        Dim Fi6 As Double = cbxFi2L2.Text
        Dim Fi7 As Double = cbxFi3L2.Text
        Dim Fi8 As Double = cbxFi4L2.Text
        Dim a5 As Double = txta1L2.Text
        Dim a6 As Double = txta2L2.Text
        Dim a7 As Double = txta3L2.Text
        Dim a8 As Double = txta4L2.Text
#End Region
#Region "vẽ trục"
        Dim pTrx1 As New cSTR_Point(P1.X, P1.Y + 450)
        Dim pTrx2 As New cSTR_Point(P1.X, P1.Y - h1 - 600)
        AddLine(pTrx1.X, pTrx1.Y, pTrx2.X, pTrx2.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTrx1 As cSTR_Line = Return_Offset_Line(pTrx1, pTrx2, L1)
        AddLine(lineOffsetTrx1.X1, lineOffsetTrx1.Y1, lineOffsetTrx1.X2, lineOffsetTrx1.Y2, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTrx2 As cSTR_Line = Return_Offset_Line(pTrx1, pTrx2, L1 + L2)
        AddLine(lineOffsetTrx2.X1, lineOffsetTrx2.Y1, lineOffsetTrx2.X2, lineOffsetTrx2.Y2, SYS_LAYER_AXIS_NAME)

        Dim pTry1 As New cSTR_Point(P1.X - 300, P1.Y)
        Dim pTry2 As New cSTR_Point(P1.X + L1 + L2 + 660, P1.Y)
        AddLine(pTry1.X, pTry1.Y, pTry2.X, pTry2.Y, SYS_LAYER_AXIS_NAME)
        Dim lineOffsetTry As cSTR_Line = Return_Offset_Line(pTry1, pTry2, -h1)
        AddLine(lineOffsetTry.X1, lineOffsetTry.Y1, lineOffsetTry.X2, lineOffsetTry.Y2, SYS_LAYER_AXIS_NAME)
#End Region
#Region " MC1"
        ' vẽ hình dưới
        Dim pNgang1 As New cSTR_Point(P1.X, P1.Y)
        Dim pNgang2 As New cSTR_Point(P1.X + L1, P1.Y)
        Dim lineNgangTr As New cSTR_Line(pNgang1.X, pNgang1.Y, pNgang2.X, pNgang2.Y)
        AddLine(pNgang1.X, pNgang1.Y, pNgang2.X, pNgang2.Y)

        Dim pNgangD1 As New cSTR_Point(P1.X, P1.Y - t)
        Dim pNgangD2 As New cSTR_Point(P1.X + 10, P1.Y - t)

        Dim lineNgangDuoi As New cSTR_Line(pNgangD1.X, pNgangD1.Y, pNgangD2.X, pNgangD2.Y)
        AddLine(lineNgangDuoi.X1, lineNgangDuoi.Y1, lineNgangDuoi.X2, lineNgangDuoi.Y2)
        'Dim lineOffsetN As cSTR_Line = Return_Offset_Line(pNgang1, pNgang2, -t)
        'AddLine(lineOffsetN.X1, lineOffsetN.Y1, lineOffsetN.X2, lineOffsetN.Y2)
        'AddLine(P1.X, P1.Y, P1.X + L1, P1.Y)
        AddLine(P1.X + L1, P1.Y, P1.X + L1, P1.Y - x)
        'AddLine(P1.X + L1, P1.Y - x, P1.X + L1 + L2, P1.Y - h1)
        AddLine(P1.X, P1.Y, P1.X, P1.Y - t)
        'OffsetLine(P1.X + L1, P1.Y - x, P1.X + L1 + L2, P1.Y - h1, -t)

        'đoạn chéo trên
        Dim pCheo1 As New cSTR_Point(P1.X + L1, P1.Y - x)
        Dim pCheo2 As New cSTR_Point(P1.X + L1 + L2, P1.Y - h1)
        AddLine(pCheo1.X, pCheo1.Y, pCheo2.X, pCheo2.Y)

        ' offset xuống khoảng t
        Dim lineOffsetC As cSTR_Line = Return_Offset_Line(pCheo1, pCheo2, -t)
        'AddLine(lineOffsetC.X1, lineOffsetC.Y1, lineOffsetC.X2, lineOffsetC.Y2)

        Dim pGiao As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineNgangDuoi, lineOffsetC)
        AddLine(lineNgangDuoi.X1, lineNgangDuoi.Y1, pGiao.X, pGiao.Y)
        Dim pGiao2 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffsetC, lineOffsetTry)
        AddLine(pGiao.X, pGiao.Y, pGiao2.X, pGiao2.Y)

        'dim
        AddDimX(P1.X, P1.X + L1, P1.Y - h1 - 600 - 50, -175)
        AddDimX(P1.X + L1, P1.X + L1 + L2, P1.Y - h1 - 600 - 50, -175)

        AddDimY(P1.X + L1 + 200, P1.Y, P1.Y - x, 175)
        AddDimY(P1.X + L1 + +L2 + 600 + 110, P1.Y, P1.Y - h1, 175)
        AddDimY(P1.X - 50, P1.Y, P1.Y - t, -175)

        'Add_CosCD(P1.X - 50 - 175 + 50 * SYS_D_DimFoot1, P1.Y, h1)
        'Dim xText As String = "%%UMẶT CẮT 1-1"
        'Add_Text_M_BIGText_with_Layer_WFactor(X0 + (L1 + L2 + L3) / 2, Y0 - Y - 300 - SYS_D_DimFoot1 - SYS_D_TextH_BIG - 200, xText, SYS_L_TEXT_TCK, 0.8)

        'thep
        'L1
        AddLine(P1.X + 35, P1.Y - t + Abv, P1.X + L1 - 35, P1.Y - t + Abv, SYS_LAYER_STEEL_NAME)
        Dim Point_Array As New ArrayList()
        Point_Array.Add(New Point2d(P1.X + Abv, P1.Y - t + Abv))        ' điểm dưới trái
        Point_Array.Add(New Point2d(P1.X + Abv, P1.Y - Abv))            ' điểm trên trái
        Point_Array.Add(New Point2d(P1.X + L1 - Abv, P1.Y - Abv))       ' điểm trên phải
        Point_Array.Add(New Point2d(P1.X + L1 - Abv, P1.Y - t - 100))   ' điểm dưới phải
        Add_PLine(Point_Array, SYS_LAYER_STEEL_NAME)

        'Phương Y (thep cham) 
        Dim a_bardot As Integer = 22
        Dim Loca_Bar3 As ArrayList
        Loca_Bar3 = Add_Bar_Dot_YL1_Tren(P1.X, P1.Y - Abv - 15, P1.X + L1, P1.Y - Abv - 15, a2, Fi2, a2, Fi4, a4, False, False, a_bardot)
        Add_Bar_Dot_YL1_Tren(P1.X, P1.Y - Abv - 15, P1.X + L1, P1.Y - Abv - 15, a2, Fi2, a2, Fi4, a4, True, True, a_bardot)

        Loca_Bar3 = Add_Bar_Dot_YL1_Duoi(P1.X, P1.Y - t + Abv + 15, P1.X + L1, P1.Y - t + Abv + 15, a1, Fi1, a1, Fi3, a3, False, False, a_bardot)
        Add_Bar_Dot_YL1_Duoi(P1.X, P1.Y - t + Abv + 15, P1.X + L1, P1.Y - t + Abv + 15, a1, Fi1, a1, Fi3, a3, True, True, a_bardot)

        'L2
        'Phương X
        Dim lineOffset_Thep_Cheo_Tren As cSTR_Line = Return_Offset_Line(pCheo1, pCheo2, -Abv)
        Dim lineNull As New cSTR_Line(P1.X + Abv, P1.Y - Abv, P1.X + L1 - Abv, P1.Y - Abv) ' đoạn ảo offset của đoạn ngang trên để tìm giao điểm
        Dim GiaoThepMoc1 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffset_Thep_Cheo_Tren, lineNull)
        Dim Giaothep As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffset_Thep_Cheo_Tren, lineOffsetTry) ' giao điểm của thép chính lớp trên với trục y-a1

        Dim ThepNgangTr As New cSTR_Line(Giaothep.X, Giaothep.Y, GiaoThepMoc1.X, GiaoThepMoc1.Y)
        AddLine(ThepNgangTr.X1, ThepNgangTr.Y1, ThepNgangTr.X2, ThepNgangTr.Y2, SYS_LAYER_STEEL_NAME)

        Dim Point_Array1 As New ArrayList()
        Point_Array1.Add(New Point2d(GiaoThepMoc1.X, GiaoThepMoc1.Y))
        Point_Array1.Add(New Point2d(GiaoThepMoc1.X - 200, GiaoThepMoc1.Y))
        Point_Array1.Add(New Point2d(GiaoThepMoc1.X - 200 + 50, GiaoThepMoc1.Y - 30))
        Add_PLine(Point_Array1, SYS_LAYER_STEEL_NAME)

        Dim lineOffset_Thep_Cheo_Duoi As cSTR_Line = Return_Offset_Line(pCheo1, pCheo2, -t + Abv)
        Dim lineNull2 As cSTR_Line = Return_Offset_Line(pNgang1, pNgang2, -Abv - 10)
        Dim GiaoThepMoc2 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffset_Thep_Cheo_Duoi, lineNull2)
        Dim GiaoThep2 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineOffset_Thep_Cheo_Duoi, lineOffsetTry) 'giao điểm của thép chính lớp dưới với trục y-a1
        AddLine(GiaoThep2.X, GiaoThep2.Y, GiaoThepMoc2.X, GiaoThepMoc2.Y, SYS_LAYER_STEEL_NAME)

        Dim Point_Array2 As New ArrayList()
        Point_Array2.Add(New Point2d(GiaoThepMoc2.X, GiaoThepMoc2.Y))
        Point_Array2.Add(New Point2d(GiaoThepMoc2.X - 200, GiaoThepMoc2.Y))
        Point_Array2.Add(New Point2d(GiaoThepMoc2.X - 200 + 50, GiaoThepMoc2.Y - 30))
        Add_PLine(Point_Array2, SYS_LAYER_STEEL_NAME)

        'Phương Y (thep cham) 

        Dim pGiao3 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(ThepNgangTr, lineNgangTr)
        Dim lineNull3_offset As cSTR_Line = Return_Offset_Line(pGiao3, Giaothep, -15)
        'Dim pGiao_Tag As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang()
        Loca_Bar3 = Add_Bar_Dot_Y_L2_Tren(lineNull3_offset.X1, lineNull3_offset.Y1, lineNull3_offset.X2, lineNull3_offset.Y2, a5, Fi5, a5, Fi5, a5, False, False, a_bardot)
        Add_Bar_Dot_Y_L2_Tren(lineNull3_offset.X1, lineNull3_offset.Y1, lineNull3_offset.X2, lineNull3_offset.Y2, a5, Fi5, a5, Fi5, a5, True, True, a_bardot)

        Dim pGiao4 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineNgangTr, lineOffset_Thep_Cheo_Duoi)
        Dim lineNull4_offset As cSTR_Line = Return_Offset_Line(pGiao4, GiaoThep2, 15)
        Loca_Bar3 = Add_Bar_Dot_Y_L2_Tren(lineNull4_offset.X1, lineNull4_offset.Y1, lineNull4_offset.X2, lineNull4_offset.Y2, a6, Fi6, a6, Fi6, a6, False, False, a_bardot)
        Add_Bar_Dot_Y_L2_Tren(lineNull4_offset.X1, lineNull4_offset.Y1, lineNull4_offset.X2, lineNull4_offset.Y2, a6, Fi6, a6, Fi6, a6, True, False, a_bardot)
#End Region
        DialogResult = Windows.Forms.DialogResult.OK

    End Sub
#End Region
    Sub veMC2(X0 As Double, Y0 As Double)
    End Sub
End Class