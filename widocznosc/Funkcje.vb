Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD

 Namespace Widocznosc
    Public Class Funkcje
       Public Function widocznosc_wymagana(ByVal spadek As Double, ByVal predkosc_ As Integer) As Double
            '                                -10   -8   -6   -4   -2   0    2    4    6    8    10
            Dim odleglosci As Integer(,) = {{390, 390, 390, 350, 330, 310, 300, 290, 280, 280, 280},    '130
                                            {340, 340, 340, 310, 290, 270, 260, 250, 240, 240, 240},    '120
                                            {280, 280, 280, 260, 240, 230, 220, 200, 200, 200, 200},    '110
                                            {220, 220, 220, 200, 180, 180, 180, 170, 170, 170, 170},    '100
                                            {190, 190, 170, 170, 150, 150, 150, 130, 130, 120, 120},    '90
                                            {160, 160, 140, 140, 120, 120, 120, 110, 110, 100, 100},    '80
                                            {110, 110, 100, 100, 90, 90, 90, 85, 85, 80, 80},           '70
                                            {80, 80, 80, 80, 70, 70, 70, 60, 60, 60, 60},               '60
                                            {55, 55, 55, 55, 50, 50, 50, 45, 45, 45, 45},               '50
                                            {40, 40, 40, 40, 35, 35, 35, 35, 35, 35, 35},               '40
                                            {25, 25, 25, 25, 20, 20, 20, 20, 20, 20, 20}}               '30

            'Dim result As String = builder.ToString()
            '      Dim result As String = predkosc(0, predkosc.GetUpperBound(1)).ToString()

            Dim kolumna As Integer
            If spadek <= -0.1 Then
                kolumna = 0
            ElseIf spadek <= -0.08 Then
                kolumna = 1
            ElseIf spadek <= -0.06 Then
                kolumna = 2
            ElseIf spadek <= -0.04 Then
                kolumna = 3
            ElseIf spadek <= -0.02 Then
                kolumna = 4
            ElseIf spadek < 0.02 Then
                kolumna = 5
            ElseIf spadek < 0.04 Then
                kolumna = 6
            ElseIf spadek < 0.06 Then
                kolumna = 7
            ElseIf spadek < 0.08 Then
                kolumna = 8
            ElseIf spadek < 0.1 Then
                kolumna = 9
            Else
                kolumna = 10
            End If


            Dim wiersz As Integer
            If predkosc_ >= 130 Then
                wiersz = 0
            ElseIf predkosc_ >= 120 Then
                wiersz = 1
            ElseIf predkosc_ >= 110 Then
                wiersz = 2
            ElseIf predkosc_ >= 100 Then
                wiersz = 3
            ElseIf predkosc_ >= 90 Then
                wiersz = 4
            ElseIf predkosc_ >= 80 Then
                wiersz = 5
            ElseIf predkosc_ >= 70 Then
                wiersz = 6
            ElseIf predkosc_ >= 60 Then
                wiersz = 7
            ElseIf predkosc_ >= 50 Then
                wiersz = 8
            ElseIf predkosc_ >= 40 Then
                wiersz = 9
            Else
                wiersz = 10
            End If

            Return odleglosci(wiersz, kolumna)

        End Function
        
        Public Function rysuj_linie(ByVal pktStart As Point3d, ByVal pktKoniec As Point3d) As Line
            Dim oDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim oDB As Database = oDWG.Database
            Dim oTrans As Transaction = oDB.TransactionManager.StartTransaction
            Dim oBT As BlockTable = CType(oDB.BlockTableId.GetObject(DatabaseServices.OpenMode.ForRead), BlockTable)
            Dim oBTR As BlockTableRecord = CType(oBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite, False, False), BlockTableRecord)
            Dim retVal As ObjectId
            Dim myLine As New Line

            Try
                myLine = New Line(pktStart, pktKoniec)

                oBTR.AppendEntity(myLine)
                oTrans.AddNewlyCreatedDBObject(myLine, True)
                oTrans.Commit()
                retVal = myLine.ObjectId
            Catch ex As System.Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Error in rysuj_linie")
            Finally
                oTrans.Dispose()
            End Try

            Return myLine
        End Function

        Public Function zrob_blok(ByVal OBidy As ObjectIdCollection, ByVal pktbloku As Point3d, ByVal nazwa_bloku As String) As Boolean
            Dim oDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim oDB As Database = oDWG.Database
            Dim oTrans As Transaction = oDB.TransactionManager.StartTransaction
            'Dim oBT As BlockTable = CType(oDB.BlockTableId.GetObject(DatabaseServices.OpenMode.ForWrite), BlockTable)
            Dim oBT As BlockTable = CType(oTrans.GetObject(oDB.BlockTableId, OpenMode.ForWrite), BlockTable)
            Dim oBTR As BlockTableRecord = New BlockTableRecord

            'Dim oLine As DatabaseServices.Line = Nothing
            Dim retVal As Boolean = False
            'Dim nazwa_bloku As String = "o" & System.Text.RegularExpressions.Regex.Replace(nazwaTXT.Text, "[\\<>/?"":;*|,=`]", "_") & "_"

            oBTR.Origin = pktbloku
            oBTR.Name = nazwa_bloku & UnixTimestamp()
            Try
                'oBT.UpgradeOpen()
                oBT.Add(oBTR)
                'oBT.DowngradeOpen()

                oTrans.AddNewlyCreatedDBObject(oBTR, True)
                oBTR.AssumeOwnershipOf(OBidy)

                Dim bref As BlockReference
                bref = New BlockReference(pktbloku, oBTR.ObjectId)
                oBTR = CType(oTrans.GetObject(oDB.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                oBTR.AppendEntity(bref)
                oTrans.AddNewlyCreatedDBObject(bref, True)

                oTrans.Commit()
            Catch ex As System.Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Błąd tworzenia bloku")
            Finally
                ' Clean up the transaction object (very important)
                oTrans.Dispose()
            End Try
            ' Return true if everything worked, false if not
            Return retVal
        End Function

        Public Function UnixTimestamp() As String
            Dim origin As DateTime = New DateTime(2010, 11, 3, 0, 0, 0, 0)
            Dim diff As TimeSpan = Date.Now - origin
            Return Hex(Math.Floor(diff.TotalSeconds))
        End Function
    End Class
End Namespace