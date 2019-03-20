Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Printing

Namespace GestionDataGridView
    Public Class PrintDataGridView
        'Donnée à imprimer
        Private WithEvents _PrintDoc As New PrintDocument
        Private _DataGV As DataGridView

        'Paramètres d'affichages
        Private Texte As [String]
        Private Font1, Font2 As Font
        Private Brush1 As SolidBrush
        Private Height, Height_ET As Single
        Private drawRect As RectangleF
        Private drawFormat As StringFormat
        Private Stylo1, Stylo2 As Pen

        'Paramètres d'impression
        Private NumLigne As Integer
        Private nbPage As Integer = 0
        Private nbColonne As Integer
        Private width As List(Of Single)
        Private Xbegin, Ybegin, Xend, Xpage As Integer
        Private XPrint As Single = 0
        Private Ratio As Single
        Dim Xb, Yb As Integer
        Private NextPage As Boolean

#Region "Configuration des paramètres"
        Private _Mode, _Style As Integer
        Public Enum Affichage As Integer
            Defaut = 0
            Selection = 1
        End Enum
        Public Enum Style As Integer
            Grille = 0
        End Enum
        Private Sub Parametre()
            Select Case _Mode
                Case 0
                    'corps du doc
                    Font1 = New Font("Arial", 12, FontStyle.Italic)
                    Stylo1 = New Pen(Color.Black)

                    'Entete
                    Font2 = New Font("Arial", 16, FontStyle.Bold)
                    Stylo2 = New Pen(Color.Black)
                    Height_ET = 50

                    'Commun
                    Brush1 = New SolidBrush(Color.Black)
                    drawFormat = New StringFormat With {.Alignment = StringAlignment.Center}
            End Select
        End Sub
#End Region

#Region "Constructeur"
        Sub New(ByRef MyDataGridView As DataGridView)
            _DataGV = MyDataGridView
        End Sub
#End Region

#Region "Procèdure externe"
        Public Overloads Sub Impression(Optional ByVal Aff As Affichage = Affichage.Defaut, Optional ByVal MonStyle As Style = Style.Grille)
            _Mode = MonStyle
            _Style = Aff
            Parametre()
            PreparePrinting()

            Dim P As New PrintPreviewDialog
            _PrintDoc.DefaultPageSettings.Landscape = True
            _PrintDoc.DefaultPageSettings.Margins = New Margins(Conv(10), Conv(10), Conv(10), Conv(10))
            P.Document = _PrintDoc
            P.ShowDialog()
        End Sub
#End Region

#Region "Processus d'impression"
        Private Sub PreparePrinting()
            'nb de colonne par ligne à imprimer
            nbColonne = _DataGV.ColumnCount
            'Numéro de la 1ere ligne imprimée
            NumLigne = 0
            'détermination de la longueur totale a imprimer (hors entete)
            For Each row As DataGridViewRow In _DataGV.Rows
                XPrint += row.Height
            Next
        End Sub
        Private Sub BuildPageToPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles _PrintDoc.PrintPage
            If nbPage = 0 Then
                'Détermination point(0,0)
                Xbegin = e.MarginBounds.Left
                Ybegin = e.MarginBounds.Top
                'Détermination point (Max,0)
                Xend = e.MarginBounds.Bottom
                'Détermination du Width en fonction de l'entete
                Ratio = RatioCol(e.MarginBounds.Width)
                width = DetermineWidth()
            End If

            'Réinitialisation Paramètres
            NextPage = False
            Xb = 0
            Yb = 0
            Xpage = Xend

            'application du style d'impression 
            Call PrintStyle(e)

            'Détermination du nombre de ligne à imprimer pour cette page
            If NextPage = False Then
                e.HasMorePages = False
                Call PreparePrinting()
            Else
                e.HasMorePages = True
            End If
            nbPage += 1
        End Sub
#End Region

#Region "Style impression"
        'Imprime en fonction de la variable _style
        Private Sub PrintStyle(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
            Dim StartPage As Boolean = True
            Do
                'coordonnée dynamique des rectangles sur l'axe X
                Xb = Xbegin


                'condition GÉNÉRALE d'arret d'impression
                On Error GoTo stoping
                'si ligne non visible, on passe
                If _DataGV.Rows(NumLigne).Visible = False Then
                    NumLigne += 1
                    Continue Do
                End If
                'condition STYLE1 d'arret d'impression
                If _Style = 1 Then
                    'si ligne non sélectionnée, on passe
                    If _DataGV.Rows(NumLigne).Selected = False Then
                        NumLigne += 1
                        Continue Do
                    End If
                End If
                On Error Resume Next


                'impression d'un ligne, boucle sur toutes les colonnes
                For i = 1 To nbColonne
                    'si colonne non visible, on passe
                    If _DataGV.Columns(i - 1).Visible = False Then Continue For

                    If StartPage Then
                        'imprime l'Entete lors de chaque nouvelle page
                        drawRect = New RectangleF(Xb, Yb, width(i - 1), Height_ET)
                        Texte = _DataGV.Columns(i - 1).HeaderText
                        e.Graphics.DrawRectangle(Stylo2, Xb, Yb, width(i - 1), Height_ET)
                        e.Graphics.DrawString(Texte, Font2, Brush1, drawRect, drawFormat)
                    Else
                        'on calcule la hauteur de la future ligne
                        Height = DetermineHeight(e, _DataGV.Rows(NumLigne), Font1, drawFormat)
                        'on détermine s'il reste de la place sur la feuille en fonction de height
                        If Xpage - Height < 0 Then
                            XPrint -= Xpage
                            If XPrint > 1 Then NextPage = True Else NextPage = False
                            Exit Do
                        End If
                        'si place pour la future ligne, on imprime
                        drawRect = New RectangleF(Xb, Yb, width(i - 1), Height)
                        'si cellule vide, on affiche rien
                        If IsDBNull(_DataGV.Rows(NumLigne).Cells(i - 1).Value) Then
                            Texte = Nothing
                        Else
                            Texte = _DataGV.Rows(NumLigne).Cells(i - 1).Value
                        End If
                        e.Graphics.DrawRectangle(Stylo1, Xb, Yb, width(i - 1), Height)
                        e.Graphics.DrawString(Texte, Font1, Brush1, drawRect, drawFormat)
                    End If
                    'coordonnée prochaine colonne de cette ligne
                    Xb += width(i - 1)
                Next

                'déduction de la place prise par la dernière ligne
                If StartPage Then
                    Xpage += -Height_ET
                    Yb += Height_ET
                Else
                    Xpage += -Height
                    NumLigne += 1
                    Yb += Height
                End If

                StartPage = False
            Loop
            Exit Sub
stoping:
            Err.Clear()
            NextPage = False
        End Sub
#End Region
        Private Function Conv(ByVal Millimetre As Integer) As Integer
            Return CInt(PrinterUnitConvert.Convert(Millimetre * 10.0R, PrinterUnit.TenthsOfAMillimeter, PrinterUnit.Display))
        End Function
        Private Function LongCol() As Single
            Dim L As Single
            For i = 1 To _DataGV.ColumnCount
                If _DataGV.Columns(i - 1).Visible Then
                    L += _DataGV.Columns(i - 1).Width
                End If
            Next
            Return L
        End Function
        Private Function RatioCol(ByVal LongPage As Single) As Single
            Return LongPage / LongCol()
        End Function
        Private Function DetermineHeight(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal Maligne As DataGridViewRow, ByVal Police As Font, ByVal Format As StringFormat) As Single
            Dim H As SizeF
            Dim Hmin As Single = 0
            For Each Cell As DataGridViewCell In Maligne.Cells
                If IsDBNull(Cell.Value) Then Continue For
                If _DataGV.Columns(Cell.ColumnIndex).Visible = False Then Continue For
                H = e.Graphics.MeasureString(Cell.Value, Police, width(Cell.ColumnIndex), Format)
                If H.Height > Hmin Then Hmin = H.Height
            Next
            Return Hmin
        End Function
        Private Function DetermineWidth() As List(Of Single)
            'détermine la largeur de la 1ere ligne du datagridview
            Dim L As New List(Of Single)
            For i = 0 To _DataGV.ColumnCount - 1
                L.Add(_DataGV.Columns(i).Width * Ratio)
            Next
            Return L
        End Function
    End Class
End Namespace

