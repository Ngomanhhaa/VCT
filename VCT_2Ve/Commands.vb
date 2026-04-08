Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Internal
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.Win32.SafeHandles

Namespace VCT_2Ve
    Public Class Commands
        <CommandMethod("XoaDoiTuong")>
        Public Sub XoaDoiTuong()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            '1. chon 1 doi tuong
            'Dim enRs As PromptEntityResult = ed.GetEntity(vbLf & "chon doi tuong")
            'If enRs.Status <> PromptStatus.OK Then
            '    Exit Sub
            'End If
            'Dim id As Object = enRs.ObjectId

            '2. chon nhieu doi tuong
            Dim slOp As New PromptSelectionOptions()
            slOp.MessageForAdding = vbLf & "chon cac doi tuong"

            Dim typeVal(0) As TypedValue
            typeVal.SetValue(New TypedValue(DxfCode.Start, "circle"), 0)

            Dim slFt As New SelectionFilter(typeVal)
            Dim slRs As PromptSelectionResult = ed.GetSelection(slOp, slFt)
            If slRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim ids = slRs.Value.GetObjectIds()

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                For Each id As ObjectId In ids
                    Dim ent As Entity = tr.GetObject(id, OpenMode.ForWrite)
                    ent.ColorIndex = 1
                Next

                tr.Commit()
            End Using
        End Sub

        <CommandMethod("DiChuyenDoiTuong")>
        Public Sub DiChuyenDoiTuong()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            '1. chon 1 doi tuong
            'Dim enRs As PromptEntityResult = ed.GetEntity(vbLf & "chon doi tuong")
            'If enRs.Status <> PromptStatus.OK Then
            '    Exit Sub
            'End If
            'Dim id As Object = enRs.ObjectId

            '2. chon nhieu doi tuong
            Dim slOp As New PromptSelectionOptions()
            slOp.MessageForAdding = vbLf & "chon cac doi tuong"

            Dim slRs As PromptSelectionResult = ed.GetSelection(slOp)
            If slRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim ids = slRs.Value.GetObjectIds()
            Dim ptRs As PromptPointResult = ed.GetPoint(vbLf & "chon diem goc")
            If ptRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim P1 = ptRs.Value
            ptRs = ed.GetPoint(vbLf & "chon diem di chuyyen")
            If slRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim P2 = ptRs.Value
            Dim vt As New Vector3d(P2.X - P1.X, P2.Y - P1.Y, 0)
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                For Each id As ObjectId In ids
                    Dim ent As Entity = tr.GetObject(id, OpenMode.ForWrite)
                    ent.TransformBy(Matrix3d.Displacement(vt))
                Next

                tr.Commit()
            End Using
        End Sub

        <CommandMethod("CopyDoiTuong")>
        Public Sub CopyDoiTuong()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            '1. chon 1 doi tuong
            'Dim enRs As PromptEntityResult = ed.GetEntity(vbLf & "chon doi tuong")
            'If enRs.Status <> PromptStatus.OK Then
            '    Exit Sub
            'End If
            'Dim id As Object = enRs.ObjectId

            '2. chon nhieu doi tuong
            Dim slOp As New PromptSelectionOptions()
            slOp.MessageForAdding = vbLf & "chon cac doi tuong"

            Dim slRs As PromptSelectionResult = ed.GetSelection(slOp)
            If slRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim ids = slRs.Value.GetObjectIds()
            Dim ptRs As PromptPointResult = ed.GetPoint(vbLf & "chon diem goc")
            If ptRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim P1 = ptRs.Value
            ptRs = ed.GetPoint(vbLf & "chon diem di chuyyen")
            If slRs.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim P2 = ptRs.Value
            Dim vt As New Vector3d(P2.X - P1.X, P2.Y - P1.Y, 0)
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim listObjectIds As New List(Of Object)
                Dim curSpace As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                For Each id As ObjectId In ids
                    Dim ent As Entity = tr.GetObject(id, OpenMode.ForWrite)
                    Dim copyEnt As Entity = ent.Clone
                    copyEnt.TransformBy(Matrix3d.Displacement(vt))
                    Dim newId As ObjectId = curSpace.AppendEntity(copyEnt)
                    tr.AddNewlyCreatedDBObject(copyEnt, True)
                    listObjectIds.Add(newId)
                Next

                'di chuyen cac doi tuong moi copy
                For Each id As ObjectId In listObjectIds
                    Dim ent As Entity = tr.GetObject(id, OpenMode.ForWrite)
                    ent.TransformBy(Matrix3d.Displacement(vt))

                Next
                tr.Commit()
            End Using
        End Sub

        <CommandMethod("CopyStyle")>
        Public Sub CopyStyle()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim destDb As Database = doc.Database
            Dim sourcedb As New Database(False, True)
            Using doc.LockDocument
                Using sourcedb
                    sourcedb.ReadDwgFile("C:\KetcauSoft\Com\KCS_STYLE.dwg", FileShare.Read, True, "")
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

        <CommandMethod("VeCauThanggg")>
        Public Sub VeCauThanggg()
            Dim frm As New frmMain
            frm.ShowDialog()
        End Sub
    End Class
End Namespace
