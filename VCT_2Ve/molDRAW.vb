Imports Autodesk
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Module molDRAW
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
    Public Sub Add_PLine(ByVal pArray As ArrayList, ByVal tLayer As String)
        Dim acDoc As Document = AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If pArray.Count = 0 Then Exit Sub
        Using LOCDOC As DocumentLock = acDoc.LockDocument
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, DatabaseServices.OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite)
                Dim acPoly As Polyline = New Polyline()
                acPoly.SetDatabaseDefaults()
                For i = 0 To pArray.Count - 1
                    Dim tPoint2D As Point2d = pArray(i)
                    acPoly.AddVertexAt(i, tPoint2D, 0, 0, 0)
                Next
                acPoly.Layer = tLayer
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                acTrans.Commit()
            End Using
        End Using
    End Sub
    Function GetArrowObjectId(ByVal newArrName As String) As ObjectId
        Dim arrObjId As ObjectId = ObjectId.Null
        Dim doc As Document = ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim oldArrName As String = TryCast(ApplicationServices.Application.GetSystemVariable("DIMBLK"), String)
        ApplicationServices.Application.SetSystemVariable("DIMBLK", newArrName)
        If oldArrName.Length <> 0 Then ApplicationServices.Application.SetSystemVariable("DIMBLK", oldArrName)
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, DatabaseServices.OpenMode.ForRead), BlockTable)
            arrObjId = bt(newArrName)
            tr.Commit()
        End Using

        Return arrObjId
    End Function
    Sub Add_Circle_ST(ByVal X As Decimal, ByVal Y As Decimal, Is_X_Direction As Boolean)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
            Dim acCirc As Circle = New Circle()
            acCirc.SetDatabaseDefaults()
            If Is_X_Direction = True Then
                acCirc.Center = New Point3d(X, Y - 62.5, 0)
            Else
                acCirc.Center = New Point3d(X + 62.5, Y, 0)
            End If
            acCirc.Radius = SYS_D_TextH_SM
            acCirc.Layer = SYS_L_SOTHEP_CIRCLE
            acCirc.ColorIndex = 1
            acBlkTblRec.AppendEntity(acCirc)
            acTrans.AddNewlyCreatedDBObject(acCirc, True)
            acTrans.Commit()
        End Using
    End Sub
    Sub Add_Text_ST(ByVal X As Decimal, ByVal Y As Decimal, ByVal tText As String, Is_X_Direction As Boolean)
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
            acText.VerticalMode = TextVerticalMode.TextVerticalMid
            If Is_X_Direction = False Then
                acText.AlignmentPoint = New Point3d(X + 62.5, Y, 0)
                acText.Rotation = Math.PI / 2
            Else
                acText.AlignmentPoint = New Point3d(X, Y - 62.5, 0)
            End If
            SET_TEXT_STYLE(acText, acTextStyleTblRec.Item(SYS_TextStyle))
            acText.Height = SYS_D_TextH_SM
            acText.TextString = tText
            acText.Layer = SYS_L_SOTHEP_TEXT
            acText.WidthFactor = SYS_D_TextWF
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            acTrans.Commit()
        End Using
    End Sub
    Sub Add_SoThep(ByVal X As Decimal, ByVal Y As Decimal, ByVal SoThep As String, Is_X_Direction As Boolean)
        Add_Circle_ST(X, Y, Is_X_Direction)
        Add_Text_ST(X, Y, SoThep, Is_X_Direction)
    End Sub
    Sub Add_QLeader(P1 As Point3d, P2 As Point3d, P3 As Point3d, ArrLdrName As String)
        Dim acDoc As Document = ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, DatabaseServices.OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite)
            Dim acDimStyleTbl As DimStyleTable
            acDimStyleTbl = acTrans.GetObject(acCurDb.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            '' Create the leader
            Dim acLdr As Leader = New Leader()
            acLdr.SetDatabaseDefaults()
            acLdr.AppendVertex(P1)
            acLdr.AppendVertex(P2)
            acLdr.AppendVertex(P3)
            acLdr.Dimldrblk = GetArrowObjectId(ArrLdrName) '"_ArchTick" "_DotBlank"
            acLdr.DimensionStyle = acDimStyleTbl.Item(SYS_DIM_STYLE)
            acLdr.Layer = SYS_L_DIM
            '' Add the new object to Model space and the transaction
            acBlkTblRec.AppendEntity(acLdr)
            acTrans.AddNewlyCreatedDBObject(acLdr, True)
            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Sub Add_QLeader(P1 As Point3d, P2 As Point3d, ArrLdrName As String)
        Dim acDoc As Document = ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, DatabaseServices.OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite)
            Dim acDimStyleTbl As DimStyleTable
            acDimStyleTbl = acTrans.GetObject(acCurDb.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            '' Create the leader
            Dim acLdr As Leader = New Leader()
            acLdr.SetDatabaseDefaults()
            acLdr.AppendVertex(P1)
            acLdr.AppendVertex(P2)
            acLdr.Dimldrblk = GetArrowObjectId(ArrLdrName) '"_ArchTick" "_DotBlank"
            acLdr.DimensionStyle = acDimStyleTbl.Item(SYS_DIM_STYLE)
            acLdr.Layer = SYS_L_DIM
            '' Add the new object to Model space and the transaction
            acBlkTblRec.AppendEntity(acLdr)
            acTrans.AddNewlyCreatedDBObject(acLdr, True)
            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Sub Add_Text_L_SMText(ByVal X As Decimal, ByVal Y As Decimal, ByVal tText As String, Is_X_Direction As Boolean)
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
            acText.Position = New Point3d(X, Y, 0)
            SET_TEXT_STYLE(acText, acTextStyleTblRec.Item(SYS_TextStyle))
            If Is_X_Direction = False Then
                acText.Rotation = Math.PI / 2
            End If
            acText.Height = SYS_D_TextH_SM
            acText.TextString = tText
            acText.WidthFactor = SYS_D_TextWF
            acText.Layer = SYS_L_TEXT
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            acTrans.Commit()
        End Using
    End Sub

    Sub Add_Text_L_SMText_3(ByVal X As Decimal, ByVal Y As Decimal, ByVal tText As String, ByVal tText_Height As Decimal, ByVal tText_Layer As String)
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
            acText.Position = New Point3d(X, Y, 0)
            SET_TEXT_STYLE(acText, acTextStyleTblRec.Item(SYS_TextStyle))
            acText.Height = tText_Height
            acText.TextString = tText
            acText.WidthFactor = SYS_D_TextWF
            acText.Layer = tText_Layer
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            acTrans.Commit()
        End Using
    End Sub

    Sub Add_Text_R_SMText(ByVal X As Decimal, ByVal Y As Decimal, ByVal tText As String, Is_X_Direction As Boolean)
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
            acText.HorizontalMode = TextHorizontalMode.TextRight
            acText.AlignmentPoint = New Point3d(X, Y, 0)
            SET_TEXT_STYLE(acText, acTextStyleTblRec.Item(SYS_TextStyle))
            If Is_X_Direction = False Then
                acText.Rotation = Math.PI / 2
            End If
            acText.Height = SYS_D_TextH_SM
            acText.TextString = tText
            acText.WidthFactor = SYS_D_TextWF
            acText.Layer = SYS_L_TEXT
            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            acTrans.Commit()
        End Using
    End Sub
    Sub Add_Bar_Tag(Bar_Dot_Point As Point3d, Bar_Long_Point As Point3d, L As Line, Bar_Dot_Text As String, Bar_Long_Text As String, Bar_Dot_Num As String, Bar_Long_Num As String, Is_Bar_Bot As Boolean)
        Dim _Bar_Dot_Point As Point3d = L.GetClosestPointTo(Bar_Dot_Point, False)
        Dim Vecto As Vector3d = New Vector3d(_Bar_Dot_Point.X - Bar_Dot_Point.X, _Bar_Dot_Point.Y - Bar_Dot_Point.Y, _Bar_Dot_Point.Z - Bar_Dot_Point.Z)
        Dim Bar_Dot_Point1 As Point3d
        Dim Bar_Dot_Point2 As Point3d
        Dim Bar_Long_Point1 As Point3d
        Dim Bar_Long_Point2 As Point3d
        If Is_Bar_Bot = True Then
            Bar_Dot_Point1 = ReturnPointByVecto(Vecto, Bar_Dot_Point, 185)
            Bar_Dot_Point2 = New Point3d(Bar_Dot_Point1.X + 350, Bar_Dot_Point1.Y, Bar_Dot_Point1.Z)
            Bar_Long_Point1 = Get_Intersect_Point(Bar_Dot_Point.X, Bar_Dot_Point.Y, Bar_Dot_Point1.X, Bar_Dot_Point1.Y,
                                                  Bar_Dot_Point1.X, Bar_Dot_Point1.Y, Bar_Dot_Point2.X, Bar_Dot_Point2.Y, -100, -125, 1)(0)
            Bar_Long_Point2 = New Point3d(Bar_Dot_Point2.X, Bar_Dot_Point2.Y - 125, Bar_Long_Point1.Z)
            Add_SoThep(Bar_Dot_Point2.X + 62.5, Bar_Dot_Point2.Y + 62.5, Bar_Dot_Num, True)
            Add_Text_R_SMText(Bar_Dot_Point2.X - 15, Bar_Dot_Point2.Y + 25, Bar_Dot_Text, True)
            Add_SoThep(Bar_Long_Point2.X + 62.5, Bar_Long_Point2.Y + 62.5, Bar_Long_Num, True)
            Add_Text_R_SMText(Bar_Long_Point2.X - 15, Bar_Long_Point2.Y + 25, Bar_Long_Text, True)
        Else
            Bar_Dot_Point1 = ReturnPointByVecto(Vecto, Bar_Dot_Point, 160)
            Bar_Dot_Point2 = New Point3d(Bar_Dot_Point1.X - 350, Bar_Dot_Point1.Y, Bar_Dot_Point1.Z)
            Bar_Long_Point1 = Get_Intersect_Point(Bar_Dot_Point.X, Bar_Dot_Point.Y, Bar_Dot_Point1.X, Bar_Dot_Point1.Y,
                                                  Bar_Dot_Point1.X, Bar_Dot_Point1.Y, Bar_Dot_Point2.X, Bar_Dot_Point2.Y, -100, -125, 1)(0)
            Bar_Long_Point2 = New Point3d(Bar_Dot_Point2.X, Bar_Dot_Point2.Y + 125, Bar_Long_Point1.Z)
            Add_SoThep(Bar_Dot_Point2.X - 62.5, Bar_Dot_Point2.Y + 62.5, Bar_Dot_Num, True)
            Add_Text_L_SMText(Bar_Dot_Point2.X + 15, Bar_Dot_Point2.Y + 25, Bar_Dot_Text, True)
            Add_SoThep(Bar_Long_Point2.X - 62.5, Bar_Long_Point2.Y + 62.5, Bar_Long_Num, True)
            Add_Text_L_SMText(Bar_Long_Point2.X + 15, Bar_Long_Point2.Y + 25, Bar_Long_Text, True)
        End If

        Add_QLeader(Bar_Dot_Point, Bar_Dot_Point1, Bar_Dot_Point2, "_DotBlank")
        Add_QLeader(Bar_Long_Point, Bar_Long_Point1, Bar_Long_Point2, "_ArchTick")
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

End Module
