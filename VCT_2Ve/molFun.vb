Imports System.IO
Imports Autodesk
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Module molFun

    Public SYS_MATH_LAM_TRON_SO As Integer = 5
    Public SYS_Rebar_Space_Min As Integer = 100
    Public SYS_Rebar_Space_Max As Integer = 300

    'Hàm khởi tạo style
    Public Sub CopyStyle(filePath As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim destDb As Database = doc.Database
        Dim sourcedb As New Database(False, True)
        Using doc.LockDocument
            Using sourcedb
                sourcedb.ReadDwgFile(filePath, FileShare.Read, True, "")
                Dim copyId As New ObjectIdCollection

                Using tr As Transaction = sourcedb.TransactionManager.StartTransaction()
                    Dim blTb As BlockTable = tr.GetObject(sourcedb.BlockTableId, OpenMode.ForRead)
                    If blTb.Has("KCS_STYLE") Then
                        copyId.Add(blTb("KCS_STYLE"))
                    End If
                    tr.Commit()
                End Using
                Dim mapping As New IdMapping()
                sourcedb.WblockCloneObjects(copyId, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, False)
            End Using
        End Using
    End Sub



    Public Sub AddblockSteelNode(cirlayer As String, hatchLayer As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim blTb As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            If blTb.Has(SYS_STEEL_NODE_BLOCK_NAME) Then
                Exit Sub
            End If
            Dim blDef As BlockTableRecord = New BlockTableRecord()
            blDef.Name = SYS_STEEL_NODE_BLOCK_NAME
            blTb.Add(blDef)
            tr.AddNewlyCreatedDBObject(blDef, True)
            tr.Commit()
        End Using

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim blTb As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim blDef As BlockTableRecord = tr.GetObject(blTb(SYS_STEEL_NODE_BLOCK_NAME), OpenMode.ForWrite)
            Dim cir As Circle = New Circle()
            cir.Center = New Point3d(0, 0, 0)
            cir.Layer = cirlayer
            cir.Radius = 10
            blDef.AppendEntity(cir)
            tr.AddNewlyCreatedDBObject(cir, True)

            Dim obCol As New ObjectIdCollection
            obCol.Add(cir.ObjectId)

            Dim hatch As New Hatch
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
            hatch.Associative = True
            hatch.AppendLoop(HatchLoopTypes.Outermost, obCol)
            hatch.EvaluateHatch(True)
            hatch.Layer = hatchLayer
            tr.AddNewlyCreatedDBObject(hatch, True)

            tr.Commit()
        End Using
    End Sub

    Public Sub AddLayer(name As String, colerIndex As Integer, lineTyle As String, plotetable As Boolean)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim ltTb As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            If ltTb.Has(lineTyle) = False Then
                Exit Sub
            End If
            Dim layerTb As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            If layerTb.Has(name) = True Then
                Exit Sub
            End If
            Dim layer As LayerTableRecord = New LayerTableRecord()
            layer.Name = name
            layer.Color = Color.FromColorIndex(ColorMethod.ByAci, colerIndex)
            layer.IsPlottable = plotetable
            layer.LinetypeObjectId = ltTb(lineTyle)
            layerTb.UpgradeOpen()
            layerTb.Add(layer)
            tr.AddNewlyCreatedDBObject(layer, True)
            tr.Commit()
        End Using

    End Sub

    Public Sub AddLText(X As Double, Y As Double, contents As String, height As Double, Optional layerName As String = SYS_LAYER_TEXT_NAME)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim textStyleTable As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim dtext As New DBText
            dtext.Position = New Point3d(X, Y, 0)
            dtext.Layer = layerName
            dtext.Height = height
            dtext.TextString = contents
            dtext.TextStyleId = textStyleTable(SYS_TEXT_STYLE_NAME)
            dtext.WidthFactor = SYS_TEXT_WIDTH_FACTOR
            curSpace.AppendEntity(dtext)
            tr.AddNewlyCreatedDBObject(dtext, True)
            tr.Commit()
        End Using
    End Sub

    Sub SET_TEXT_STYLE(ByVal acText As Object, ByVal TextTypeTable As Object)
        Try
            acText.TextStyle = TextTypeTable
        Catch ex As System.Exception
            'MsgBox(ex.Message)
            acText.TextStyleID = TextTypeTable
        End Try
    End Sub
    Sub Add_Text_M_BIGText_with_Layer_WFactor(ByVal X As Decimal, ByVal Y As Decimal, ByVal tText As String, ByVal tLayer As String, ByVal WFactor As Decimal)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
            Dim acTextStyleTblRec As TextStyleTable
            acTextStyleTblRec = acTrans.GetObject(acCurDb.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acText As DBText = New DBText()
            acText.SetDatabaseDefaults()
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            acText.AlignmentPoint = New Point3d(X, Y, 0)
            SET_TEXT_STYLE(acText, acTextStyleTblRec.Item(SYS_TextStyle))
            acText.Height = SYS_D_TextH_BIG
            acText.TextString = tText
            acText.WidthFactor = WFactor
            acText.Layer = tLayer
            acText.ColorIndex = 4
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            acTrans.Commit()
        End Using
    End Sub
    Public Sub AddCText(X As Double, Y As Double, contents As String, height As Double, Optional layerName As String = SYS_LAYER_TEXT_NAME)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim textStyleTable As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim dtext As New DBText
            dtext.HorizontalMode = TextHorizontalMode.TextCenter
            dtext.AlignmentPoint = New Point3d(X, Y, 0)
            dtext.Layer = layerName
            dtext.Height = height
            dtext.TextString = contents
            dtext.TextStyleId = textStyleTable(SYS_TEXT_STYLE_NAME)
            dtext.WidthFactor = SYS_TEXT_WIDTH_FACTOR
            curSpace.AppendEntity(dtext)
            tr.AddNewlyCreatedDBObject(dtext, True)
            tr.Commit()
        End Using
    End Sub

    Public Sub AddMCText(X As Double, Y As Double, contents As String, height As Double, Optional layerName As String = SYS_LAYER_TEXT_NAME)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim textStyleTable As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim dtext As New DBText
            dtext.HorizontalMode = TextHorizontalMode.TextCenter
            dtext.VerticalMode = TextVerticalMode.TextVerticalMid
            dtext.AlignmentPoint = New Point3d(X, Y, 0)
            dtext.Layer = layerName
            dtext.Height = height
            dtext.TextString = contents
            dtext.TextStyleId = textStyleTable(SYS_TEXT_STYLE_NAME)
            dtext.WidthFactor = SYS_TEXT_WIDTH_FACTOR
            curSpace.AppendEntity(dtext)
            tr.AddNewlyCreatedDBObject(dtext, True)
            tr.Commit()
        End Using
    End Sub


    Public Sub Add_BreakLineY(X As Double, tY1 As Double, tY2 As Double, layer As String)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
            Dim Y1, Y2
            Y1 = Math.Min(tY1, tY2)
            Y2 = Math.Max(tY1, tY2)
            Dim objPLine As Polyline = New Polyline
            Dim Y3 As Decimal = (Y1 + Y2) / 2
            Dim DX As Decimal = 25
            Dim DY As Decimal = 25
            objPLine.SetDatabaseDefaults()
            objPLine.AddVertexAt(0, New Point2d(X, Y1 - DX), 0, 0, 0)
            objPLine.AddVertexAt(1, New Point2d(X, Y3 - DY / 2), 0, 0, 0)
            objPLine.AddVertexAt(2, New Point2d(X - DX, Y3 - DY / 2), 0, 0, 0)
            objPLine.AddVertexAt(3, New Point2d(X + DX, Y3 + DY / 2), 0, 0, 0)
            objPLine.AddVertexAt(4, New Point2d(X, Y3 + DY / 2), 0, 0, 0)
            objPLine.AddVertexAt(5, New Point2d(X, Y2 + DX), 0, 0, 0)
            objPLine.Layer = layer
            objPLine.ColorIndex = 1
            acBlkTblRec.AppendEntity(objPLine)
            acTrans.AddNewlyCreatedDBObject(objPLine, True)
            acTrans.Commit()
        End Using
    End Sub

    'Public Sub AddTagThepChu(X As Double, Y As Double, sh As String, info As String, listPointRebar As List(Of Point2d))
    '    AddCircle(X - SYS_TAG_CIRCLE_DIA / 2 * 25, Y, SYS_TAG_CIRCLE_DIA / 2 * 25, SYS_LAYER_THIN_NAME)
    '    AddMCText(X - SYS_TAG_CIRCLE_DIA / 2 * 25, Y, sh, SYS_TEXT_HEIGHT * 25)

    '    For i = 0 To listPointRebar.Count - 1
    '        Dim pt As Point2d = listPointRebar(i)
    '        AddLine(pt.X, pt.Y, pt.X - 62.5, Y, SYS_LAYER_THIN_NAME)
    '    Next
    '    AddLine(X, Y, listPointRebar.Last.X - 62.5, Y, SYS_LAYER_THIN_NAME)

    '    AddLText(X + 25, Y + 20, info, SYS_TEXT_HEIGHT * 25)

    'End Sub

    Public Class cSTR_Point
        Public Property id As Integer
        Public Property model_id As Integer 'ID của drawingObject, Dùng khi tạo model, để lấy ngược lại UniqueName
        Public Property label As String 'uniquename
        Public Property X As Decimal
        Public Property Y As Decimal
        Public Property Z As Decimal

        Public Sub New()
        End Sub

        Public Sub New(x As Decimal, y As Decimal, Optional z As Decimal = 0)
            Me.X = x
            Me.Y = y
            Me.Z = z
        End Sub
    End Class

    Public Class cSTR_Line
        Public Sub New(x1 As Decimal, y1 As Decimal, x2 As Decimal, y2 As Decimal)
            Me.X1 = x1
            Me.Y1 = y1
            Me.X2 = x2
            Me.Y2 = y2
        End Sub

        Public Property id As Integer
        Public Property X1 As Decimal
        Public Property Y1 As Decimal
        Public Property X2 As Decimal
        Public Property Y2 As Decimal
        Public Property Get_P1 As cSTR_Point
        Public Property Get_P2 As cSTR_Point
    End Class



    Public Function Return_Giao_Diem_Hai_Doan_Thang(Line1 As cSTR_Line, Line2 As cSTR_Line) As cSTR_Point
        Dim xPoint As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(Line1.X1, Line1.Y1, Line1.X2, Line1.Y2, Line2.X1, Line2.Y1, Line2.X2, Line2.Y2)
        Return xPoint
    End Function

    Public Function Return_Giao_Diem_Hai_Doan_Thang(
        x1 As Decimal, y1 As Decimal,
        x2 As Decimal, y2 As Decimal,
        x3 As Decimal, y3 As Decimal,
        x4 As Decimal, y4 As Decimal) As cSTR_Point

        Dim denominator As Decimal = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)

        If denominator = 0 Then
            Return Nothing
        End If

        Dim px As Decimal =
            ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / denominator

        Dim py As Decimal =
            ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / denominator

        'xác đinh 2 đoạn thẳng có cắt nhau không, nếu không thì kéo dài và trả về giao diểm
        Dim PointOnSeg1 As Boolean = PointOnSegment(px, py, x1, y1, x2, y2)
        Dim PointOnSeg2 As Boolean = PointOnSegment(px, py, x3, y3, x4, y4)

        If PointOnSeg1 <> PointOnSeg2 Then
            Return New cSTR_Point(px, py, 0)
        End If
        Return New cSTR_Point(px, py, 0)

        Return Nothing
    End Function

    'xác định phạm vi đoạn thẳng
    Private Function PointOnSegment(
        px As Decimal, py As Decimal,
        x1 As Decimal, y1 As Decimal,
        x2 As Decimal, y2 As Decimal) As Boolean

        Dim minX As Decimal = Math.Min(x1, x2)
        Dim maxX As Decimal = Math.Max(x1, x2)
        Dim minY As Decimal = Math.Min(y1, y2)
        Dim maxY As Decimal = Math.Max(y1, y2)
        Dim tol As Decimal = 0.000001

        If px < minX - tol OrElse px > maxX + tol Then Return False
        If py < minY - tol OrElse py > maxY + tol Then Return False

        Return True
    End Function


    Function Return_Offset_Point(p1 As cSTR_Point, p2 As cSTR_Point, pSource As cSTR_Point, delta As Decimal) As cSTR_Point
        ' Compute angle from p1 to p2 (radians)
        Dim vx As Double = CDbl(p2.X - p1.X)
        Dim vy As Double = CDbl(p2.Y - p1.Y)
        Dim anfa As Double = Math.Atan2(vy, vx)
        Dim anfa1 As Double = anfa + Math.PI / 2.0

        Dim x As Decimal = pSource.X + CDec(delta * Math.Cos(anfa1))
        Dim y As Decimal = pSource.Y + CDec(delta * Math.Sin(anfa1))

        x = Math.Round(x, SYS_MATH_LAM_TRON_SO)
        y = Math.Round(y, SYS_MATH_LAM_TRON_SO)

        Dim pt As New cSTR_Point()
        pt.X = x
        pt.Y = y
        pt.Z = 0
        Return pt
    End Function

    Function Return_Offset_Line(P1 As cSTR_Point, P2 As cSTR_Point, offset_value As Decimal) As cSTR_Line
        ' Offset the line to the left (relative to P1->P2)
        Dim P1A As cSTR_Point = Return_Offset_Point(P1, P2, P1, offset_value)
        Dim P2A As cSTR_Point = Return_Offset_Point(P1, P2, P2, offset_value)
        Return New cSTR_Line(P1A.X, P1A.Y, P2A.X, P2A.Y)
    End Function

    Function ReturnPointByVecto(vecto As Vector3d, P1 As Point3d, distance As Double) As Point3d
        Dim n As Double = distance / vecto.Length
        Dim X As Double = vecto.X * n
        Dim Y As Double = vecto.Y * n
        Dim Z As Double = vecto.Z * n
        Dim P2 As Point3d = New Point3d(X + P1.X, Y + P1.Y, Z + P1.Z)
        Return P2
    End Function
    Function Get_Intersect_Point(X1a As Decimal, Y1a As Decimal, X1b As Decimal, Y1b As Decimal, X2a As Decimal, Y2a As Decimal, X2b As Decimal, Y2b As Decimal, Curves1 As Decimal, Curves2 As Decimal, Type As Integer) _
        As Point3dCollection
        Dim D1 As Line = Get_OffSet_Line(X1a, Y1a, X1b, Y1b, Curves1)(0) 'đường thẳng offset 1
        Dim D2 As Line
        If Type = 1 Then
            D2 = Get_OffSet_Line(X2a, Y2a, X2b, Y2b, Curves2)(0) 'đường thẳng offset 2
        ElseIf Type = 2 Then
            D2 = New Line(New Point3d(X2a, Y2a, 0), New Point3d(X2b, Y2b, 0)) 'Đường thẳng từ 2 điểm 
        End If
        Dim P1 As New Point3dCollection
        D1.IntersectWith(D2, Intersect.ExtendBoth, P1, IntPtr.Zero, IntPtr.Zero)
        Return P1
    End Function

    Public Sub Add_SNode(ByVal X As Decimal, ByVal Y As Decimal, ByVal Draw_Scale As Decimal)
        Dim acDoc As Document = AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite)
            Dim InsertID As ObjectId = acBlkTbl.Item("KCS_SNODE")
            Dim BlockRef As BlockReference = New BlockReference(New Point3d(X, Y, 0), InsertID)
            BlockRef.Layer = SYS_LAYER_STEEL_NAME
            BlockRef.ScaleFactors = New Scale3d(Draw_Scale / 25, Draw_Scale / 25, Draw_Scale / 25)
            acBlkTblRec.AppendEntity(BlockRef)
            acTrans.AddNewlyCreatedDBObject(BlockRef, True)
            acTrans.Commit()
        End Using
    End Sub

    Sub Add_SteelTop(X1 As Decimal, Y1 As Decimal, X2 As Decimal, Y2 As Decimal, Is_X_Direction As Boolean)
        Dim Point_Array As ArrayList = New ArrayList()
        Select Case Is_X_Direction
            Case True
                Dim P1 As Point2d = New Point2d(X1, Y1 - 75)
                Point_Array.Add(P1)
                P1 = New Point2d(X1, Y1)
                Point_Array.Add(P1)
                P1 = New Point2d(X2, Y2)
                Point_Array.Add(P1)
                P1 = New Point2d(X2, Y2 - 75)
                Point_Array.Add(P1)
            Case False
                Dim P1 As Point2d = New Point2d(X1 + 75, Y1)
                Point_Array.Add(P1)
                P1 = New Point2d(X1, Y1)
                Point_Array.Add(P1)
                P1 = New Point2d(X2, Y2)
                Point_Array.Add(P1)
                P1 = New Point2d(X2 + 75, Y2)
                Point_Array.Add(P1)
        End Select
        Add_PLine(Point_Array, SYS_LAYER_STEEL_NAME)
    End Sub

    Function Get_OffSet_Line(X1 As Decimal, Y1 As Decimal, X2 As Decimal, Y2 As Decimal, Curves As Decimal) As DBObjectCollection
        Dim acDbObjColl As DBObjectCollection
        Dim acLine1 As Line = New Line(New Point3d(X1, Y1, 0), New Point3d(X2, Y2, 0))
        acDbObjColl = acLine1.GetOffsetCurves(Curves)

        Return acDbObjColl
    End Function
    Sub Add_SNode(ByVal X As Decimal, ByVal Y As Decimal)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite)
            Dim InsertID As ObjectId = acBlkTbl.Item("KCS_SNODE")
            Dim BlockRef As BlockReference = New BlockReference(New Point3d(X, Y, 0), InsertID)
            BlockRef.Layer = SYS_LAYER_STEEL_NAME
            acBlkTblRec.AppendEntity(BlockRef)
            acTrans.AddNewlyCreatedDBObject(BlockRef, True)
            acTrans.Commit()
        End Using
    End Sub
    Function Add_Bar_Dot(X1 As Decimal, Y1 As Decimal, X2 As Decimal, Y2 As Decimal, A As Decimal, Is_bot As Boolean, Optional a_bardot As Integer = 15) As ArrayList
        Dim L As Line = New Line(New Point3d(X1, Y1, 0), New Point3d(X2, Y2, 0))
        Dim Length As Decimal = L.Length - 104
        Dim N As Decimal = Length \ A
        Dim Center_SNode As Decimal = (Length - A * N) / 2
        Dim P1 As Point3d
        Dim P2 As Point3d
        Dim L1 As Line

        For i As Integer = 0 To N
            Add_SNode(L.GetPointAtDist(52 + i * A + Center_SNode).X, L.GetPointAtDist(52 + i * A + Center_SNode).Y)
            If i = N \ 3 Then
                If Is_bot = True Then
                    L1 = Get_OffSet_Line(X1, Y1, X2, Y2, -a_bardot)(0)
                    P1 = L.GetPointAtDist(52 + i * A + Center_SNode)
                    P2 = L1.GetPointAtDist(52 + i * A + Center_SNode - 100)
                Else
                    L1 = Get_OffSet_Line(X1, Y1, X2, Y2, a_bardot)(0)
                    P1 = L.GetPointAtDist(52 + i * A + Center_SNode)
                    P2 = L1.GetPointAtDist(52 + i * A + Center_SNode + 100)
                End If
            End If
        Next
        Return New ArrayList({P1, P2})
    End Function
    Function Return_Bar_Notes_N_Fi(ByVal N As Integer, ByVal Fi As Integer) As String
        Dim tValue As String = N & SYS_KCS_CONFIG_BarNote_KyHieuDuongKinh & Fi
        Return tValue
    End Function

    Function Return_Bar_Notes_Fi_A(ByVal Fi As Integer, ByVal A As Integer) As String
        Dim tValue As String = SYS_KCS_CONFIG_BarNote_KyHieuDuongKinh & Fi & SYS_KCS_CONFIG_BarNote_KyHieuKhoangCach & A
        Return tValue
    End Function
    Function Get_Length_TK(P1 As Point3d, P2 As Point3d) As Decimal
        Dim L As Decimal = Math.Round(P1.DistanceTo(P2), 0)
        If L Mod 10 <> 0 Then
            L = L + (10 - L Mod 10)
        End If
        Return L
    End Function
    Function Get_Length_TK(Length As Decimal) As Decimal
        Dim L As Decimal = Math.Round(Length, 0)
        If L Mod 50 <> 0 Then
            L = L + (50 - L Mod 50)
        End If
        Return L
    End Function
End Module


