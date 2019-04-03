Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Printing

Namespace GestionDataGridView
    Public Class PrintDataGridView
        Private Const ERR_IMPRESSION = "erreur lors du processus d'impression"
        Private Const ERR_NODATA = "Le dataGridView est vides, il ne contient aucune donnée (pour le rendre visible vous devez l'ajouter à un control Form selon la méthode [monForm.Controls.Add(monDGV)])"

        'Donnée à imprimer
        Private WithEvents _PrintDoc As New PrintDocument
        Private _DataGV As DataGridView

        'Paramètres d'affichages ENTETE
        Private _Entete_Font As Font = New Font("Arial", 10, FontStyle.Bold)
        Private _Entete_Stylo As Pen = New Pen(Color.Black)
        Private _Entete_Brush As SolidBrush = New SolidBrush(Color.Black)
        Private _Entete_StringFormat As StringFormat = New StringFormat With {.Alignment = StringAlignment.Center,
                                                                              .LineAlignment = StringAlignment.Center}
        'Paramètres d'affichages CORPS DU DOCUMENT
        Private _CorpsTexte_Font As Font = New Font("Arial", 10, FontStyle.Regular)
        Private _CorpsTexte_Stylo As Pen = New Pen(Color.Black)
        Private _CorpsTexte_Brush As SolidBrush = New SolidBrush(Color.Black)
        Private _CorpsTexte_StringFormat As StringFormat = New StringFormat With {.Alignment = StringAlignment.Center}

        'Paramètres d'affichages Ligne individuelle
        Private _LigneIndiv_Font As Font = New Font("Arial", 10, FontStyle.Italic)
        Private _LigneIndiv_Stylo As Pen = New Pen(Color.White)
        Private _LigneIndiv_Brush As SolidBrush = New SolidBrush(Color.Black)
        Private _LigneIndiv_StringFormat As StringFormat = New StringFormat With {.Alignment = StringAlignment.Near}
        Private _LigneIndiv_Text As String = Nothing

        'Coordonnées de la zone d'impression
        Private _Xbegin, _Ybegin, _Xend, _Yend As Integer
        Private _LigneEnCours As Integer = 1
        Private _modeAffichage As Affichage

        'Détermine la position d'écriture de la nouvelle ligne
        Private _Ycursor As Integer

        'nombre de page imprimées
        Private _nbPage As Integer = 0

        'Détermine si une nouvelle page est nécessaire pour finir l'impression
        Private _NextPage As Boolean

#Region "Property"
        Public Property textePremierePage As String
            Get
                Return _LigneIndiv_Text
            End Get
            Set(ByVal value As String)
                _LigneIndiv_Text = value
            End Set
        End Property
        Public Property Font_Entete As Font
            Get
                Return _Entete_Font
            End Get
            Set(ByVal value As Font)
                _Entete_Font = value
            End Set
        End Property
        Public Property Font_CorpsTexte As Font
            Get
                Return _CorpsTexte_Font
            End Get
            Set(ByVal value As Font)
                _CorpsTexte_Font = value
            End Set
        End Property
        Public Property Font_LigneIndiv As Font
            Get
                Return _LigneIndiv_Font
            End Get
            Set(ByVal value As Font)
                _LigneIndiv_Font = value
            End Set
        End Property
        Public Property CouleurEncadrement_Entete As Pen
            Get
                Return _Entete_Stylo
            End Get
            Set(ByVal value As Pen)
                _Entete_Stylo = value
            End Set
        End Property
        Public Property CouleurEncadrement_CorpsTexte As Pen
            Get
                Return _CorpsTexte_Stylo
            End Get
            Set(ByVal value As Pen)
                _CorpsTexte_Stylo = value
            End Set
        End Property
        Public Property CouleurEncadrement_LigneIndiv As Pen
            Get
                Return _LigneIndiv_Stylo
            End Get
            Set(ByVal value As Pen)
                _LigneIndiv_Stylo = value
            End Set
        End Property
        Public Property CouleurTexte_Entete As SolidBrush
            Get
                Return _Entete_Brush
            End Get
            Set(ByVal value As SolidBrush)
                _Entete_Brush = value
            End Set
        End Property
        Public Property CouleurTexte_CorpsTexte As SolidBrush
            Get
                Return _CorpsTexte_Brush
            End Get
            Set(ByVal value As SolidBrush)
                _CorpsTexte_Brush = value
            End Set
        End Property
        Public Property CouleurTexte_LigneIndiv As SolidBrush
            Get
                Return _LigneIndiv_Brush
            End Get
            Set(ByVal value As SolidBrush)
                _LigneIndiv_Brush = value
            End Set
        End Property
#End Region

#Region "Enumération"
        Public Enum Affichage As Byte
            All = 1
            Selection = 2
        End Enum
#End Region

#Region "Constructeur"
        Sub New(ByRef MyDataGridView As DataGridView)
            If MyDataGridView.Rows.Count = 0 Then
                Throw New VBToolsException(ERR_NODATA)
            End If
            _DataGV = MyDataGridView
        End Sub
#End Region

#Region "Procèdure externe"
        Public Overloads Sub Impression(Optional ByVal mode As Affichage = Affichage.All)
            _modeAffichage = mode

            Try
                Dim P As New PrintPreviewDialog
                _PrintDoc.DefaultPageSettings.Landscape = True
                _PrintDoc.DefaultPageSettings.Margins = New Margins(Conv(10), Conv(10), Conv(10), Conv(10))
                P.Document = _PrintDoc
                P.ShowDialog()
            Catch ex As Exception
                Throw New VBToolsException(ERR_IMPRESSION, ex)
            End Try
        End Sub
#End Region

#Region "Processus d'impression"
        Private Sub DeterminationParametreImpression(ByRef e As System.Drawing.Printing.PrintPageEventArgs)
            'Détermination point (0,0)
            _Xbegin = e.MarginBounds.Left
            _Ybegin = e.MarginBounds.Top
            'Détermination point (Max,Max)
            _Xend = e.MarginBounds.Right
            _Yend = e.MarginBounds.Bottom
        End Sub

        ''' <summary>
        ''' Se déclenche pour chaques pages
        ''' </summary>
        Private Sub ConstructionImpressionPages(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles _PrintDoc.PrintPage
            '--------------------------------
            '-  initialisation du document  -
            '--------------------------------
            If _nbPage = 0 Then
                'détermination de la zone imprimable
                DeterminationParametreImpression(e)
            End If

            '-------------------------------
            '-  Initialisation de la page  -
            '-------------------------------
            'on commence en haut du document
            _Ycursor = _Ybegin
            _NextPage = False


            '--------------------------
            '-  Ecriture ligne indiv  -
            '--------------------------
            If _nbPage = 0 Then
                If Not IsNothing(_LigneIndiv_Text) Then
                    WriteLine(e, _LigneIndiv_Text)
                End If
            End If


            '--------------------------
            '-  Ecriture de l'entete  -
            '--------------------------
            WriteDataGridViewHeader(e)


            '-------------------------------
            '-  Ecriture du Corps de page  -
            '-------------------------------
            For i = _LigneEnCours To _DataGV.Rows.Count
                'si ligne non sélectionnée en mode sélection
                If _modeAffichage = Affichage.All Or (_modeAffichage = Affichage.Selection And _DataGV.Rows(i - 1).Selected) Then
                    WriteDataGridViewRow(e, i - 1)
                End If
                If _NextPage Then
                    Exit For
                End If
            Next

            'Détermination du nombre de ligne à imprimer pour cette page
            e.HasMorePages = _NextPage

            'fin de la page et incrémentation
            _nbPage += 1
        End Sub

        ''' <summary>
        ''' Ecriture d'une ligne
        ''' </summary>
        Private Sub WriteLine(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal Texte As String)
            Dim monRectangle As Rectangle
            Dim coeffLigne As Integer = determineLignes(Texte)

            'on prepare le format de la cellule
            monRectangle = New Rectangle(_Xbegin, _Ycursor, _Xend - _Xbegin, _LigneIndiv_Font.GetHeight() * coeffLigne * 1.05)

            'on complète l'intérieur
            e.Graphics.DrawString(Texte, _LigneIndiv_Font, _LigneIndiv_Brush, monRectangle, _LigneIndiv_StringFormat)
            'on dessine le contour de la cellule
            e.Graphics.DrawRectangle(_LigneIndiv_Stylo, Rectangle.Round(monRectangle))

            'on déplace le curseur
            _Ycursor += _LigneIndiv_Font.GetHeight() * coeffLigne * 1.05
        End Sub

        ''' <summary>
        ''' Ecriture d'une ligne du DataGridView
        ''' </summary>
        Private Sub WriteDataGridViewRow(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal ligneDataGridView As Integer)
            Dim monRectangle As Rectangle
            Dim monTexte As String
            Dim Xcursor As Single = _Xbegin
            Dim Xcol As Single = 0
            Dim coeffLigne As Integer = determineLignes(ligneDataGridView)

            'si ligne à écrire sort de la zone d'impression
            If _Ycursor + (_CorpsTexte_Font.GetHeight() * coeffLigne) > _Yend Then
                _LigneEnCours = ligneDataGridView
                _NextPage = True
                Exit Sub
            End If

            For i = 1 To _DataGV.ColumnCount
                If _DataGV.Columns(i - 1).Visible = True Then

                    'calcul de la longueur du rectangle en proportion équivalente a celle du DataGridView
                    Xcol = _DataGV.Columns(i - 1).Width * (_Xend - _Xbegin) / XTotalDataGridView()

                    'on prepare le format de la cellule
                    monRectangle = New Rectangle(Xcursor, _Ycursor, Xcol, _CorpsTexte_Font.GetHeight() * coeffLigne * 1.05)
                    'le texte de la cellule
                    monTexte = _DataGV.Rows(ligneDataGridView).Cells(i - 1).Value

                    'on complète l'intérieur
                    e.Graphics.DrawString(monTexte, _CorpsTexte_Font, _CorpsTexte_Brush, monRectangle, _CorpsTexte_StringFormat)
                    'on dessine le contour de la cellule
                    e.Graphics.DrawRectangle(_CorpsTexte_Stylo, Rectangle.Round(monRectangle))
                    'prochaine coordonnée d'écriture
                    Xcursor += Xcol

                End If
            Next

            'on déplace le curseur
            _Ycursor += _CorpsTexte_Font.GetHeight() * coeffLigne * 1.05
        End Sub

        ''' <summary>
        ''' Ecriture de l'entete du DataGridView
        ''' </summary>
        Private Sub WriteDataGridViewHeader(ByRef e As System.Drawing.Printing.PrintPageEventArgs)
            Dim monRectangle As RectangleF
            Dim monTexte As String
            Dim Xcursor As Single = _Xbegin
            Dim Xcol As Single = 0
            Dim coeffLigne As Integer = determineLignes()

            For i = 1 To _DataGV.ColumnCount
                If _DataGV.Columns(i - 1).Visible = True Then
                    'calcul de la longueur du rectangle en proportion équivalente a celle du DataGridView
                    Xcol = _DataGV.Columns(i - 1).Width * (_Xend - _Xbegin) / XTotalDataGridView()
                    'on prepare le format de la cellule
                    monRectangle = New RectangleF(Xcursor, _Ycursor, Xcol, _Entete_Font.GetHeight() * coeffLigne * 1.05)
                    'le texte de la cellule
                    monTexte = _DataGV.Columns(i - 1).HeaderText
                    'on dessine le contour de la cellule
                    e.Graphics.DrawRectangle(_Entete_Stylo, Rectangle.Round(monRectangle))
                    'on complète l'intérieur
                    e.Graphics.DrawString(monTexte, _Entete_Font, _Entete_Brush, monRectangle, _Entete_StringFormat)
                    'prochaine coordonnée d'écriture
                    Xcursor += Xcol
                End If
            Next

            'on déplace le curseur
            _Ycursor += _Entete_Font.GetHeight() * coeffLigne * 1.05
        End Sub
#End Region

#Region "Fonction interne"
        Private Function Conv(ByVal Millimetre As Integer) As Integer
            Return CInt(PrinterUnitConvert.Convert(Millimetre * 10.0R, PrinterUnit.TenthsOfAMillimeter, PrinterUnit.Display))
        End Function
        Private Function XTotalDataGridView() As Single
            Dim L As Single
            For i = 1 To _DataGV.ColumnCount
                If _DataGV.Columns(i - 1).Visible Then
                    L += _DataGV.Columns(i - 1).Width
                End If
            Next
            Return L
        End Function
        ''' <summary>
        ''' Retourne le nombre de caractère max affichable par une ligne
        ''' </summary>
        ''' <returns>le nombre de caractère max</returns>
        ''' <remarks>
        ''' formule déterminé par excel : 
        ''' nbChar = 549.16 * police^-1.039
        ''' r² = 0.9996
        ''' pour une longueur de colonne de 364
        ''' police 40 --> 12 chars
        ''' police 20 --> 24 chars
        ''' police 10 --> 51 chars
        ''' police 8 --> 63 chars
        ''' </remarks>
        Private Function nbCharMaxParLigne(ByVal police As Single, ByVal tailleColonne As Single) As Integer
            Dim nbCar As Single = Math.Pow(police, -1.039) * 549.16
            Return tailleColonne * nbCar / 364
        End Function
        ''' <summary>
        ''' Retourne le nombre de lignes nécessaires à l'affichage
        ''' </summary>
        ''' <param name="numeroLigneDGV">numéro de ligne du DGV (-1 pour l'entete)</param>
        ''' <returns>nombre de ligne affichée</returns>
        ''' <remarks></remarks>
        Private Function determineLignes(Optional ByVal numeroLigneDGV As Integer = -1)
            Dim coeffMultiplicateur As Integer = 1
            Dim Xcol As Single = 0


            For i = 1 To _DataGV.ColumnCount
                Dim coeff As Integer
                Dim texte As String

                Xcol = _DataGV.Columns(i - 1).Width * (_Xend - _Xbegin) / XTotalDataGridView()

                If numeroLigneDGV = -1 Then
                    texte = _DataGV.Columns(i - 1).HeaderText
                    coeff = (texte.Length / nbCharMaxParLigne(_Entete_Font.Size, Xcol)) + 1
                Else
                    texte = _DataGV.Rows(numeroLigneDGV).Cells(i - 1).Value
                    coeff = (texte.Length / nbCharMaxParLigne(_CorpsTexte_Font.Size, Xcol)) + 1
                End If

                If coeff > coeffMultiplicateur Then coeffMultiplicateur = coeff
            Next

            Return coeffMultiplicateur
        End Function
        Private Function determineLignes(ByVal monTexte As String)
            Return (monTexte.Length / nbCharMaxParLigne(_LigneIndiv_Font.Size, _Xend - _Xbegin)) + 1
        End Function
#End Region
    End Class
End Namespace

