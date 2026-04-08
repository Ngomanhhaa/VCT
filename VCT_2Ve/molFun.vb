Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Module molFun
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

    Class cSTR_Point
        Public Property id As Integer
        Public Property model_id As Integer 'ID của drawingObject, Dùng khi tạo model, để lấy ngược lại UniqueName
        Public Property label As String 'uniquename
        Public Property X As Decimal
        Public Property Y As Decimal
        Public Property Z As Decimal
    End Class

    Public Class cSTR_Line
        Public Property X1 As cSTR_Line
        Public Property Y1 As cSTR_Line
        Public Property X2 As cSTR_Line
        Public Property Y2 As cSTR_Line

    End Class

    Function Return_Giao_Diem_Hai_Doan_Thang(Line1 As cSTR_Line, Line2 As cSTR_Line) As cSTR_Point
        Dim xPoint As cSTR_Point = Return_Giao_Diem_Hai_Doan_Thang(Line1.X1, Line1.Y1, Line1.X2, Line1.Y2, Line2.X1, Line2.Y1, Line2.X2, Line2.Y2)
        Return xPoint
    End Function


End Module

