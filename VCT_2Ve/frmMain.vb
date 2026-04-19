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

        Dim L1 As Double = txtL1.Text
        Dim L2 As Double = txtL2.Text
        Dim L3 As Double = txtL3.Text
        Dim h1 As Double = txtH1.Text * 1000
        Dim h2 As Double = txtH2.Text * 1000
        Dim t As Double = txtT.Text
        Dim x As Double = txtX.Text
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

        Dim P1b, P2a, P2b, P3a, P3b As New Point3dCollection 'Intersect Point Border
        've truc
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
        AddDimY(P1.X + L1 + +L2 + 600 + 200, P1.Y, P1.Y - h1, 175)
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
        Loca_Bar3 = Add_Bar_Dot(P1.X, P1.Y - Abv - 15, P1.X + L1, P1.Y - Abv - 15, a1, False, a_bardot)
        Add_Bar_Dot(P1.X, P1.Y - Abv - 15, P1.X + L1, P1.Y - Abv - 15, a1, True, a_bardot)

        Loca_Bar3 = Add_Bar_Dot(P1.X, P1.Y - t + Abv + 15, P1.X + L1, P1.Y - t + Abv + 15, a2, False, a_bardot)
        Add_Bar_Dot(P1.X, P1.Y - t + Abv + 15, P1.X + L1, P1.Y - t + Abv + 15, a2, True, a_bardot)

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
        Loca_Bar3 = Add_Bar_Dot(lineNull3_offset.X1, lineNull3_offset.Y1, lineNull3_offset.X2, lineNull3_offset.Y2, a5, False, a_bardot)
        Add_Bar_Dot(lineNull3_offset.X1, lineNull3_offset.Y1, lineNull3_offset.X2, lineNull3_offset.Y2, a6, True, a_bardot)

        Dim pGiao4 As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(lineNgangTr, lineOffset_Thep_Cheo_Duoi)
        Dim lineNull4_offset As cSTR_Line = Return_Offset_Line(pGiao4, GiaoThep2, 15)
        Loca_Bar3 = Add_Bar_Dot(lineNull4_offset.X1, lineNull4_offset.Y1, lineNull4_offset.X2, lineNull4_offset.Y2, a6, False, a_bardot)
        Add_Bar_Dot(lineNull4_offset.X1, lineNull4_offset.Y1, lineNull4_offset.X2, lineNull4_offset.Y2, a6, True, a_bardot)


        Dim Bar1_X_Bot As STR_CONFIG_BAR = ThangBo_Config.Bar1_X_Bot
        Dim Bar1_X_Top As STR_CONFIG_BAR = ThangBo_Config.Bar1_X_Top
        Dim Bar1_Y_Bot As STR_CONFIG_BAR = ThangBo_Config.Bar1_Y_Bot
        Dim Bar1_Y_Top As STR_CONFIG_BAR = ThangBo_Config.Bar1_Y_Top
        Dim Bar3_X_Bot As STR_CONFIG_BAR = ThangBo_Config.Bar3_X_Bot
        Dim Bar3_X_Top As STR_CONFIG_BAR = ThangBo_Config.Bar3_X_Top
        Dim Bar3_Y_Bot As STR_CONFIG_BAR = ThangBo_Config.Bar3_Y_Bot
        Dim Bar3_Y_Top As STR_CONFIG_BAR = ThangBo_Config.Bar3_Y_Top

        Dim bar_num_even As Integer = 2
        Dim bar_num_odd As Integer = 1
        Dim Pdb1, Pdb2 As New Point3dCollection 'Point Dot Bar 1 and 2
        Dim Location_Bar_Bot_Tag_L2 As ArrayList
        Location_Bar_Bot_Tag_L2 = Add_Bar_Dot(Pdb1(0).X, Pdb1(0).Y, Pdb2(0).X, Pdb2(0).Y, Bar1_X_Bot.A, True) 'Add Bar Dot Bot L2
        Dim Bot_L2 As Line = New Line(P2a(0), P2b(0))
        Dim Bar_Bot_Dot_L2_Text As String = SYS_KCS_CONFIG_BarNote_KyHieuDuongKinh & Bar1_X_Bot.Fi & SYS_KCS_CONFIG_BarNote_KyHieuKhoangCach & Bar1_X_Bot.A
        Dim Bar_Bot_Long_L2_Text As String = SYS_KCS_CONFIG_BarNote_KyHieuDuongKinh & Bar1_Y_Bot.Fi & SYS_KCS_CONFIG_BarNote_KyHieuKhoangCach & Bar1_Y_Bot.A
        Add_Bar_Tag(Location_Bar_Bot_Tag_L2(0), Location_Bar_Bot_Tag_L2(1), Bot_L2, Bar_Bot_Dot_L2_Text, Bar_Bot_Long_L2_Text, bar_num_even, bar_num_odd, True)
        TK_Dot_L2 = Get_Length_TK(Pdb1(0), Pdb2(0))

        DialogResult = Windows.Forms.DialogResult.OK

    End Sub

End Class