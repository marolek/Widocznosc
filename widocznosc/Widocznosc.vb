Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD

Namespace Widocznosc

    Public Class KlasaWidocznosc
        Implements IExtensionApplication
        Dim IDyObjektow As ObjectIdCollection = New ObjectIdCollection
        Private tm As TransientManager = TransientManager.CurrentTransientManager
        Private line As Line
        Dim trasa2d As Curve
        Dim interwal As Double = 5


        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor()
            ed.WriteMessage(vbNewLine + vbNewLine + "Wczytano dodatek Widoczność (autor: Mariusz Walczyński)" + vbNewLine + "Użyj polecenia 'widocznosc'")
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
            Console.WriteLine("Cleaning up...")
        End Sub

         <CommandMethod("widocznosc")>
        Public Sub GetPointsFromUser()
            Dim funkcja As Funkcje = New Funkcje()
            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            '        Dim PktObwiedni As Point3dCollection = New Point3dCollection()

            Dim entOpts As PromptEntityOptions
            entOpts = New PromptEntityOptions(vbLf & "Wybierz polilinię lub: ")

            Dim trasa As Curve
            Dim StalaOdl As Boolean = False
            Dim OdlWidocz As Double
            Dim predkosc As Integer = 0

            With entOpts
                .SetRejectMessage(vbLf & "Obiekt musi być typu 'linia/polilinia/polilinia3D/splajn/łuk/okrąg'")
                .AddAllowedClass(GetType(DatabaseServices.Polyline), False)
                .AddAllowedClass(GetType(Polyline2d), False)
                .AddAllowedClass(GetType(Polyline3d), False)
                .AddAllowedClass(GetType(Spline), False)
                .AddAllowedClass(GetType(Circle), False)
                .AddAllowedClass(GetType(Line), False)
                .AddAllowedClass(GetType(Arc), False)
                .Keywords.Add("Interwał")
                .AllowNone = True
                .AllowObjectOnLockedLayer = True
            End With

            Dim entRes As PromptEntityResult = acDoc.Editor.GetEntity(entOpts)
            If entRes.StringResult = "Interwał" Then
                'acDoc.Editor.WriteMessage("cieciwa lub trasa")
                Dim pytInterwalOpts As PromptDistanceOptions
                pytInterwalOpts = New PromptDistanceOptions(vbLf & "Podaj interwał: ")

                With pytInterwalOpts
                    .DefaultValue = interwal
                    .UseDefaultValue = True
                End With

                Dim pInterRes As PromptDoubleResult
                pInterRes = acDoc.Editor.GetDistance(pytInterwalOpts)

                interwal = pInterRes.Value 'wielkość interwału badanej odległości
                'ZmienOdstep()
                Exit Sub
            End If
            If entRes.Status <> PromptStatus.OK Then GoTo koniec

            Dim widOpcje As PromptIntegerOptions
            widOpcje = New PromptIntegerOptions(vbLf & "Podaj prędkość lub: ")

            With widOpcje
                .Keywords.Add("Odległość")
                .AllowNone = True
            End With

            Dim pIntRes As PromptIntegerResult
            pIntRes = acDoc.Editor.GetInteger(widOpcje)

            If pIntRes.StringResult = "Odległość" Then

                Dim pytOdleglOpts As PromptDistanceOptions
                pytOdleglOpts = New PromptDistanceOptions(vbLf & "Podaj odległość (stała): ")

                Dim pytOdlegRes As PromptDoubleResult
                pytOdlegRes = acDoc.Editor.GetDistance(pytOdleglOpts)

                OdlWidocz = pytOdlegRes.Value 'odległość widoczności na zatrzymanie

                'jeżeli StalaOdl = true to nie obliczaj odległości
                StalaOdl = True
                GoTo BezPredkosci

            End If

            If pIntRes.Status <> PromptStatus.OK Then GoTo koniec

            predkosc = pIntRes.Value 'prędkość projektowana

    BezPredkosci:

            Dim PlanXY As Plane = New Plane(Point3d.Origin, Vector3d.ZAxis)

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                'linia wskazujaca punkt na trasie
                line = New Line
                line.Color = Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1)
                tm.AddTransient(line, TransientDrawingMode.Main, 128, New IntegerCollection)

                Dim dbObj As DBObject = acTrans.GetObject(entRes.ObjectId, OpenMode.ForWrite)
                Dim acBlkTbl As BlockTable
                acBlkTbl = CType(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)

                Dim obwiednia As DatabaseServices.Polyline = New DatabaseServices.Polyline

                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = CType(acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                            OpenMode.ForWrite), BlockTableRecord)

                'Dim trasa As Curve = dbObj
                trasa = TryCast(dbObj, Curve)

                trasa2d = trasa.GetOrthoProjectedCurve(PlanXY)
                'acBlkTblRec.AppendEntity(trasa2D)
                'acTrans.AddNewlyCreatedDBObject(trasa2D, True)


                'trasa.Highlight()

                'dlugosc trasy
                Dim Dlugosc As Double = trasa2D.GetDistanceAtParameter(trasa2D.EndParam)

                Dim pozycja As Double = 0

                AddHandler acDoc.Editor.PointMonitor, AddressOf TrackMouse
                'wskazanie zakresu
                Dim optPointStart As New PromptPointOptions(vbCrLf + "Wskaż początek zakresu lub <Pomiń>: ")

                With optPointStart
                    .Keywords.Add("Pomiń")
                    .Keywords.Default = "Pomiń"
                    .AllowNone = True
                End With

                Dim resPointStart As PromptPointResult = acDoc.Editor.GetPoint(optPointStart)

                If resPointStart.StringResult = "Pomiń" Then GoTo pomin_poczatek

                If resPointStart.Status = PromptStatus.OK Then
                    Try
                        Dim pikietaStart As Double = trasa2D.GetDistAtPoint(trasa2D.GetClosestPointTo(resPointStart.Value, False))
                        acDoc.Editor.WriteMessage(vbCrLf + "Początek zakresu: " + pikietaStart.ToString("N2"))
                        pozycja = pikietaStart
                    Catch ex As System.Exception
                        MsgBox(ex.Message, MsgBoxStyle.Information, "Jakiś błąd")
                    End Try

                End If
    pomin_poczatek:

                Dim optPointEnd As New PromptPointOptions(vbCrLf + "Koniec zakresu: ")

                With optPointEnd
                    .Keywords.Add("Pomiń")
                    .Keywords.Default = "Pomiń"
                    .AllowNone = True
                End With

                Dim resPointEnd As PromptPointResult = acDoc.Editor.GetPoint(optPointEnd)

                If resPointEnd.StringResult = "Pomiń" Then GoTo pomin_koniec

                If resPointEnd.Status = PromptStatus.OK Then
                    Try
                        Dim pikietaEnd As Double = trasa2D.GetDistAtPoint(trasa2D.GetClosestPointTo(resPointEnd.Value, False))
                        acDoc.Editor.WriteMessage(vbCrLf + "Koniec zakresu: " + pikietaEnd.ToString("N2"))
                        Dlugosc = pikietaEnd
                    Catch ex As System.Exception
                        MsgBox(ex.Message, MsgBoxStyle.Information, "Jakiś błąd")
                    End Try

                End If
    pomin_koniec:
                'koniec

                Using acDoc.LockDocument

                    '         PktObwiedni.Add(trasa.StartPoint)
                    Dim werteks As Integer = 0
                    obwiednia.AddVertexAt(werteks, trasa2D.GetPointAtDist(pozycja).Convert2d(PlanXY), 0, 0, 0)
                    werteks = werteks + 1

                    Dim IdCieciwy As ObjectId

                    acDoc.Editor.WriteMessage(vbCrLf + "Pikieta od: " + pozycja.ToString("N2") + vbCrLf + "Pikieta do: " + Dlugosc.ToString("N2"))

                    If Not StalaOdl Then OdlWidocz = funkcja.widocznosc_wymagana(0, predkosc)

                    acDoc.Editor.WriteMessage(vbCrLf + "Widocznosc na: " + OdlWidocz.ToString("N2"))

                    Do While pozycja + OdlWidocz < Dlugosc 'cLen

                        Dim pkt1 As Point3d = trasa2D.GetPointAtDist(pozycja)
                        Dim pkt1_3d As Point3d = trasa.GetPointAtDist(pozycja) 'punkt milimetr dalej do obliczenia spadku
                        Dim pkt2_3d As Point3d = trasa.GetPointAtDist(pozycja + 0.001)
                        Dim spadek As Double = (pkt2_3d.Z - pkt1_3d.Z) / 0.001

                        'acDoc.Editor.WriteMessage(vbCrLf + String.Format("Spadek {0}", spadek))
                        If Not StalaOdl Then OdlWidocz = funkcja.widocznosc_wymagana(spadek, predkosc)

                        Dim pts3D As Point3dCollection = New Point3dCollection()

                        Using acCirc As Circle = New Circle()
                            acCirc.Center = pkt1
                            acCirc.Radius = OdlWidocz

                            ''' NIE DODAJE OKREGU DO BLOK RECORDu (potrzebne tylko do przeciecia)
                            'acBlkTblRec.AppendEntity(acCirc)
                            'acTrans.AddNewlyCreatedDBObject(acCirc, True)
                            trasa2D.IntersectWith(acCirc, Intersect.OnBothOperands, pts3D, IntPtr.Zero, IntPtr.Zero)
                            If (pts3D.Count = 0) Then
                                Exit Do
                            End If
                        End Using

                        Dim cieciwa As Line = funkcja.rysuj_linie(pkt1, IIf(trasa2D.GetDistAtPoint(pts3D(0)) > pozycja, pts3D(0), pts3D(1)))

                        IDyObjektow.Add(cieciwa.ObjectId)  'rysuj_linie zwraca ObjectID -> kolekcja IDyObjektow

                        'Jeżeli jest więcej niz jedna linia to znajdz przecięcia
                        If Not IdCieciwy.IsNull Then
                            Dim linia1 As Line = TryCast(acTrans.GetObject(IdCieciwy, OpenMode.ForRead), Line)
                            Dim pkt3D As Point3dCollection = New Point3dCollection()
                            cieciwa.IntersectWith(linia1, Intersect.OnBothOperands, pkt3D, IntPtr.Zero, IntPtr.Zero)
                            If (pkt3D.Count > 0) Then
                                obwiednia.AddVertexAt(werteks, pkt3D(0).Convert2d(PlanXY), 0, 0, 0)
                                werteks = werteks + 1
                            End If
                        End If

                        'rysuj_linie(pkt1, pkt2)
                        pozycja = pozycja + interwal
                        IdCieciwy = cieciwa.ObjectId
                    Loop
                    'PktObwiedni.Add(trasa.EndPoint)

                    obwiednia.AddVertexAt(werteks, trasa2D.GetPointAtDist(Dlugosc).Convert2d(PlanXY), 0, 0, 0)
                    acBlkTblRec.AppendEntity(obwiednia)
                    acTrans.AddNewlyCreatedDBObject(obwiednia, True)

                    'acDoc.Editor.WriteMessage(PktObwiedni.Count.ToString)
                    
                    funkcja.zrob_blok(IDyObjektow, New Point3d(0, 0, 0), IIf(StalaOdl, String.Format("Widocznosc_{0}_", OdlWidocz), String.Format("Widocznosc_V{0}_", predkosc)))
                    IDyObjektow.Clear()

                End Using  'unlockDocument
                RemoveHandler acDoc.Editor.PointMonitor, AddressOf TrackMouse
                tm.EraseTransient(line, New IntegerCollection)
                acTrans.Commit()

            End Using
    koniec:
        End Sub

        Private Sub TrackMouse(sender As Object, e As PointMonitorEventArgs)
            Dim mousePoint As Point3d = e.Context.RawPoint
            Dim pol As Point3d = trasa2d.GetClosestPointTo(mousePoint, False)
            line.StartPoint = mousePoint
            line.EndPoint = pol
            Try
                Dim station As Double = trasa2d.GetDistAtPoint(trasa2d.GetClosestPointTo(pol, False))
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbCr + "Pikieta: " + station.ToString("N2"))
                e.AppendToolTipText("Pikieta: " + station.ToString("N2"))
            Catch ex As System.Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Jakiś błąd")
            End Try
            tm.UpdateTransient(line, New IntegerCollection)
        End Sub

    End Class
End Namespace
