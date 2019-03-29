Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Printing

Namespace GestionDataGridView
    Public Class PrintDataGridView
        'Donnée à imprimer
        Private WithEvents _PrintDoc As New PrintDocument
        Private _DataGV As DataGridView

        'Paramètres d'affichages
        Private _Texte As [String]
        Private _Font_CorpsTexte, _Font_Entete As Font
        Private _Brush1 As SolidBrush
        Private _Height, _Height_Entete As Single
        Private _drawRect As RectangleF
        Private _drawFormat As StringFormat
        Private _Stylo_CorpsTexte, _Stylo_Entete As Pen

        'Paramètres d'impression
        Private _NumLigne As Integer
        Private _nbPage As Integer = 0
        Private _nbColonne As Integer
        Private _width As List(Of Single)
        Private _Xbegin, _Ybegin, _Xend, _Yend, _Xpage As Integer
        Private _XPrint As Single = 0
        Private _Ratio As Single
        Private _Xb, _Yb As Integer
        Private _NextPage As Boolean
        Private _Mode, _Style As Integer

#Region "Configuration des paramètres"
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
                    _Font_CorpsTexte = New Font("Arial", 12, FontStyle.Italic)
                    _Stylo_CorpsTexte = New Pen(Color.Black)

                    'Entete
                    _Font_Entete = New Font("Arial", 16, FontStyle.Bold)
                    _Stylo_Entete = New Pen(Color.Black)
                    _Height_Entete = 50

                    'Commun
                    _Brush1 = New SolidBrush(Color.Black)
                    _drawFormat = New StringFormat With {.Alignment = StringAlignment.Center}
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
            _nbColonne = _DataGV.ColumnCount
            'Numéro de la 1ere ligne imprimée
            _NumLigne = 0
            'détermination de la longueur totale a imprimer (hors entete)
            For Each row As DataGridViewRow In _DataGV.Rows
                _XPrint += row.Height
            Next
        End Sub
        Private Sub BuildPageToPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles _PrintDoc.PrintPage
            If _nbPage = 0 Then
                'Détermination point(0,0)
                _Xbegin = e.MarginBounds.Left
                _Ybegin = e.MarginBounds.Top
                'Détermination point (Max,0)
                _Xend = e.MarginBounds.Bottom
                _Yend = e.MarginBounds.Right
                'Détermination du Width en fonction de l'entete
                _Ratio = RatioCol(e.MarginBounds.Width)
                _width = DetermineWidth()
            End If

            'Réinitialisation Paramètres
            _NextPage = False
            _Xb = 0
            _Yb = 0
            '_Yb = 40
            _Xpage = _Xend

            'application du style d'impression 
            Call PrintStyle(e)

            'Détermination du nombre de ligne à imprimer pour cette page
            If _NextPage = False Then
                e.HasMorePages = False
                Call PreparePrinting()
            Else
                e.HasMorePages = True
            End If
            _nbPage += 1
        End Sub
#End Region

#Region "Style impression"
        'Imprime en fonction de la variable _style
        Private Sub PrintStyle(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
            Dim StartPage As Boolean = True
            Do
                'coordonnée dynamique des rectangles sur l'axe X
                _Xb = _Xbegin


                'condition GÉNÉRALE d'arret d'impression
                On Error GoTo stoping
                'si ligne non visible, on passe
                If _DataGV.Rows(_NumLigne).Visible = False Then
                    _NumLigne += 1
                    Continue Do
                End If
                'condition STYLE1 d'arret d'impression
                If _Style = 1 Then
                    'si ligne non sélectionnée, on passe
                    If _DataGV.Rows(_NumLigne).Selected = False Then
                        _NumLigne += 1
                        Continue Do
                    End If
                End If
                On Error Resume Next


                'impression d'un ligne, boucle sur toutes les colonnes
                For i = 1 To _nbColonne
                    'si colonne non visible, on passe
                    If _DataGV.Columns(i - 1).Visible = False Then Continue For

                    If StartPage Then
                        'e.Graphics.DrawString("ceci est une phrase simple d'entete", _Font_Entete, _Brush1, New PointF(_Xbegin, 0))
                        'imprime l'Entete lors de chaque nouvelle page
                        _drawRect = New RectangleF(_Xb, _Yb, _width(i - 1), _Height_Entete)
                        _Texte = _DataGV.Columns(i - 1).HeaderText
                        e.Graphics.DrawRectangle(_Stylo_Entete, _Xb, _Yb, _width(i - 1), _Height_Entete)
                        e.Graphics.DrawString(_Texte, _Font_Entete, _Brush1, _drawRect, _drawFormat)
                    Else
                        'on calcule la hauteur de la future ligne
                        _Height = DetermineHeight(e, _DataGV.Rows(_NumLigne), _Font_CorpsTexte, _drawFormat)
                        'on détermine s'il reste de la place sur la feuille en fonction de height
                        If _Xpage - _Height < 0 Then
                            _XPrint -= _Xpage
                            If _XPrint > 1 Then _NextPage = True Else _NextPage = False
                            Exit Do
                        End If
                        'si place pour la future ligne, on imprime
                        _drawRect = New RectangleF(_Xb, _Yb, _width(i - 1), _Height)
                        'si cellule vide, on affiche rien
                        If IsDBNull(_DataGV.Rows(_NumLigne).Cells(i - 1).Value) Then
                            _Texte = Nothing
                        Else
                            _Texte = _DataGV.Rows(_NumLigne).Cells(i - 1).Value
                        End If
                        e.Graphics.DrawRectangle(_Stylo_CorpsTexte, _Xb, _Yb, _width(i - 1), _Height)
                        e.Graphics.DrawString(_Texte, _Font_CorpsTexte, _Brush1, _drawRect, _drawFormat)
                    End If
                    'coordonnée prochaine colonne de cette ligne
                    _Xb += _width(i - 1)
                Next

                'déduction de la place prise par la dernière ligne
                If StartPage Then
                    _Xpage += -_Height_Entete
                    _Yb += _Height_Entete
                Else
                    _Xpage += -_Height
                    _NumLigne += 1
                    _Yb += _Height
                End If

                StartPage = False
            Loop
            Exit Sub
stoping:
            Err.Clear()
            _NextPage = False
        End Sub
#End Region

#Region "Fonction interne"
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
                H = e.Graphics.MeasureString(Cell.Value, Police, _width(Cell.ColumnIndex), Format)
                If H.Height > Hmin Then Hmin = H.Height
            Next
            Return Hmin
        End Function
        Private Function DetermineWidth() As List(Of Single)
            'détermine la largeur de la 1ere ligne du datagridview
            Dim L As New List(Of Single)
            For i = 0 To _DataGV.ColumnCount - 1
                L.Add(_DataGV.Columns(i).Width * _Ratio)
            Next
            Return L
        End Function
#End Region
    End Class
End Namespace

