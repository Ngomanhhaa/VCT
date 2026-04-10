Imports System.IO
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


    Public Sub AddLine(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Optional layerName As String = SYS_LAYER_BORDER_NAME)
        '1. khai báo các biến đại diện cho document, database và editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        '2. sử dụng transaction để tương tác các đối tượng trong database ( nên sử dụng khối lệnh using)
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            '3. khai báo các biến đại diện cho bảng dữ liệu muốn làm việc
            Dim blTb As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim modelSpace As BlockTableRecord = tr.GetObject(blTb(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            '4. khai báo đối tượng cần lấy ra hoặc thêm vào từ dâtabase (hình học, layer,..)
            Dim _line As New Line()

            Dim P1 As Point3d = New Point3d(X1, Y1, 0)
            Dim P2 As Point3d = New Point3d(X2, Y2, 0)

            _line.StartPoint = P1
            _line.EndPoint = P2
            _line.Layer = layerName
            '5. thêm vào database
            modelSpace.AppendEntity(_line)
            tr.AddNewlyCreatedDBObject(_line, True)

            '6. kết thúc giao dịch ( Commit, Abort)
            tr.Commit()

        End Using
    End Sub

    Public Sub OffsetLine(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, offsetDist As Double, Optional layerName As String = SYS_LAYER_BORDER_NAME)
        '1. khai báo các biến đại diện cho document, database và editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        '2. sử dụng transaction để tương tác các đối tượng trong database ( nên sử dụng khối lệnh using)
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            '3. khai báo các biến đại diện cho bảng dữ liệu muốn làm việc
            Dim blTb As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim modelSpace As BlockTableRecord = tr.GetObject(blTb(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            '4. khai báo đối tượng cần lấy ra hoặc thêm vào từ dâtabase (hình học, layer,..)
            Dim dx As Double = X2 - X1
            Dim dy As Double = Y2 - Y1
            Dim L As Double = Math.Sqrt(dx * dx + dy * dy)

            If L = 0 Then Exit Sub

            ' vector pháp tuyến đơn vị
            Dim nx As Double = -dy / L
            Dim ny As Double = dx / L

            ' tạo line offset
            Dim X1_off As Double = X1 + nx * offsetDist
            Dim Y1_off As Double = Y1 + ny * offsetDist
            Dim X2_off As Double = X2 + nx * offsetDist
            Dim Y2_off As Double = Y2 + ny * offsetDist

            AddLine(X1_off, Y1_off, X2_off, Y2_off, layerName)

            '6. kết thúc giao dịch ( Commit, Abort)
            tr.Commit()
        End Using
    End Sub

    Public Sub AddStirrup(botLeftPoint As Point3d, topRightPoint As Point3d, layer As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim X1, Y1, X2, Y2 As Double
        X1 = botLeftPoint.X
        Y1 = botLeftPoint.Y
        X2 = topRightPoint.X
        Y2 = topRightPoint.Y
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim blTb As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim modelSpace As BlockTableRecord = tr.GetObject(blTb(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim pline As New Polyline
            pline.SetDatabaseDefaults()

            pline.AddVertexAt(0, New Point2d(X1, Y1 + 30), 0, 0, 0)
            pline.AddVertexAt(1, New Point2d(X1 + 30, Y1), 0, 0, 0)
            pline.AddVertexAt(2, New Point2d(X2 - 30, Y1), 0, 0, 0)
            pline.AddVertexAt(3, New Point2d(X2, Y1 + 30), 0, 0, 0)
            pline.AddVertexAt(4, New Point2d(X2, Y2 - 30), 0, 0, 0)
            pline.AddVertexAt(5, New Point2d(X2 - 30, Y2), 0, 0, 0)
            pline.AddVertexAt(6, New Point2d(X1 + 30, Y2), 0, 0, 0)
            pline.AddVertexAt(7, New Point2d(X1, Y2 - 30), 0, 0, 0)
            pline.AddVertexAt(8, New Point2d(X1, Y1 + 30), 0, 0, 0)

            pline.SetBulgeAt(0, 0.4142)
            pline.SetBulgeAt(2, 0.4142)
            pline.SetBulgeAt(4, 0.4142)
            pline.SetBulgeAt(6, 0.4142)

            pline.Layer = layer

            modelSpace.AppendEntity(pline)
            tr.AddNewlyCreatedDBObject(pline, True)
            tr.Commit()
        End Using
    End Sub

    Public Sub AddSteelNode(X As Double, Y As Double, Optional layerName As String = SYS_LAYER_STEEL_NAME)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim blRf As New BlockReference(New Point3d(X, Y, 0), bt(SYS_STEEL_NODE_BLOCK_NAME))
            blRf.Layer = layerName
            curSpace.AppendEntity(blRf)
            tr.AddNewlyCreatedDBObject(blRf, True)
            tr.Commit()
        End Using
    End Sub
    Public Sub AddCircle(X As Double, Y As Double, R As Double, layer As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim blTb As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim cir As New Circle
            cir.Center = New Point3d(X, Y, 0)
            cir.Radius = R
            cir.Layer = layer

            curSpace.AppendEntity(cir)
            tr.AddNewlyCreatedDBObject(cir, True)
            tr.Commit()
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

    Public Sub AddDimX(X1 As Double, X2 As Double, Y As Double, DY As Double)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If X1 = X2 Then Exit Sub
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
            Dim acDimStyleTbl As DimStyleTable
            acDimStyleTbl = acTrans.GetObject(acCurDb.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim objDim As AlignedDimension = New AlignedDimension
            objDim.SetDatabaseDefaults()
            objDim.XLine1Point = New Point3d(X1, Y, 0)
            objDim.XLine2Point = New Point3d(X2, Y, 0)
            objDim.DimLinePoint = New Point3d(X2, Y + DY, 0)
            objDim.DimensionStyle = acDimStyleTbl(SYS_DIM_STYLE_NAME)
            objDim.Layer = SYS_LAYER_DIM_NAME
            acBlkTblRec.AppendEntity(objDim)
            acTrans.AddNewlyCreatedDBObject(objDim, True)
            acTrans.Commit()
        End Using
    End Sub

    Public Sub AddDimY(ByVal X, ByVal Y1, ByVal Y2, ByVal DX)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If Y1 = Y2 Then Exit Sub
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
            Dim acDimStyleTbl As DimStyleTable
            acDimStyleTbl = acTrans.GetObject(acCurDb.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim objDim As AlignedDimension = New AlignedDimension
            objDim.SetDatabaseDefaults()
            objDim.XLine1Point = New Point3d(X, Y1, 0)
            objDim.XLine2Point = New Point3d(X, Y2, 0)
            objDim.DimLinePoint = New Point3d(X + DX, Y1, 0)
            objDim.DimensionStyle = acDimStyleTbl(SYS_DIM_STYLE_NAME)
            objDim.Layer = SYS_LAYER_DIM_NAME
            acBlkTblRec.AppendEntity(objDim)
            acTrans.AddNewlyCreatedDBObject(objDim, True)
            acTrans.Commit()
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

    Sub Add_CosCD(ByVal X As Decimal, ByVal Y As Decimal, ByVal EL As Decimal)
        Add_CosCD_Symbol(X, Y)
        AddLine(X + 87.5, Y, X + 87.5, Y + 250, SYS_LAYER_THIN_NAME)
        AddLine(X + 87.5 - 75, Y + 115, X + 87.5 - 75 + 350, Y + 115, SYS_LAYER_THIN_NAME)

        Dim tText As String = Str(Math.Round(EL, 3)).Trim
        If EL >= 0 Then
            If EL >= 1 Then
                tText = "+" & tText
            Else
                If Left(tText, 1) = "." Then tText = "+0" & tText
            End If
            If Left(Right(tText, 2), 1) = "." Then tText = tText & "00"
            If Left(Right(tText, 3), 1) = "." Then tText = tText & "0"
            If Left(Right(tText, 4), 1) <> "." Then tText = tText & ".000"
        Else
            If EL > -1 Then
                If Left(tText, 2) = "-." Then tText = "-0" & Right(tText, Len(tText) - 1)
            End If
            If Left(Right(tText, 2), 1) = "." Then tText = tText & "00"
            If Left(Right(tText, 3), 1) = "." Then tText = tText & "0"
            If Left(Right(tText, 4), 1) <> "." Then tText = tText & ".000"
        End If
        If EL = 0 Then tText = "%%p" & tText
        AddLText(X + 87.5 - 75 + 100, Y + 115 + 20, tText, 62.5)
    End Sub

    Sub Add_CosCD_Symbol(ByVal X As Decimal, ByVal Y As Decimal)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim InsertID As ObjectId = acBlkTbl.Item("COS_CD")
            Dim BlockRef As BlockReference = New BlockReference(New Point3d(X, Y, 0), InsertID)
            BlockRef.Layer = SYS_LAYER_THIN_NAME
            acBlkTblRec.AppendEntity(BlockRef)
            acTrans.AddNewlyCreatedDBObject(BlockRef, True)
            acTrans.Commit()
        End Using
    End Sub

    Public Sub AddTagThepChu(X As Double, Y As Double, sh As String, info As String, listPointRebar As List(Of Point2d))
        AddCircle(X - SYS_TAG_CIRCLE_DIA / 2 * 25, Y, SYS_TAG_CIRCLE_DIA / 2 * 25, SYS_LAYER_THIN_NAME)
        AddMCText(X - SYS_TAG_CIRCLE_DIA / 2 * 25, Y, sh, SYS_TEXT_HEIGHT * 25)

        For i = 0 To listPointRebar.Count - 1
            Dim pt As Point2d = listPointRebar(i)
            AddLine(pt.X, pt.Y, pt.X - 62.5, Y, SYS_LAYER_THIN_NAME)
        Next
        AddLine(X, Y, listPointRebar.Last.X - 62.5, Y, SYS_LAYER_THIN_NAME)

        AddLText(X + 25, Y + 20, info, SYS_TEXT_HEIGHT * 25)

    End Sub
End Module


