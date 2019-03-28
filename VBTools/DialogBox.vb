Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing

Namespace DialogBox

    '-----------------------------------------------
    '   Classe : CHOICEBOX
    '-----------------------------------------------
    Public Class ChoiceBox
        'Liste message d'erreur
        Private Const ERR_CREATE_WINFORM = "Erreur lors de la construction de la fenêtre de dialogue"
        Private Const ERR_DOSSIER_INCONNU = "Le chemin absolu {0} n'existe pas ou n'est pas accessible"
        Private Const ERR_CREATIONCONTROLE = "Ereur lors de la création d'un control dans la fenetre principale"
        Private Const ERR_DESCRIPTION = "Ereur lors du formatage de la zone Texte de description"
        Private Const ERR_SHOWDIALOG = "Ereur lors de l'affichage de la boite de dialogue"
        Private Const ERR_CHARGETREEVIEW = "Ereur lors de la récupération des noms de fichier pour remplir la boite de dialogue"

        'extention autorisé à l'affichage
        Private Const _EXT_ALL = "*.*"

        Private _RepertoireBaseAbsolu As String
        Private _FichierSelectionne As String = Nothing
        Private _VoirFichierDansTreeView As Boolean

        Private _TextDescription As String = Nothing
        Private _TextFont As Font = Nothing
        Private _TextColor As System.Drawing.Color

        'Element du windows form
        Private WithEvents _Form_Main As New Form
        Private WithEvents _Form_Main_BT_OK As New Button
        Private WithEvents _Form_Main_BT_Annuler As New Button
        Private WithEvents _Form_Main_TreeView As New TreeView
        Private _Form_Main_TXT_Result As New TextBox
        Private _Form_Main_TXT_Description As New TextBox

#Region "Property"
        Protected WriteOnly Property setRepertoireBaseAbsolu As String
            Set(ByVal value As String)
                _RepertoireBaseAbsolu = value
            End Set
        End Property
        Protected WriteOnly Property setFichierSelectionne As String
            Set(ByVal value As String)
                _FichierSelectionne = value
            End Set
        End Property
        Public ReadOnly Property getRepertoireBaseAbsolu As String
            Get
                Return _RepertoireBaseAbsolu
            End Get
        End Property
        Public ReadOnly Property getRepertoireBaseRelatif As String
            Get
                Return Path.GetFileName(_RepertoireBaseAbsolu)
            End Get
        End Property
        Public ReadOnly Property getRepertoireBeforeBase As String
            Get
                Dim frac() = _RepertoireBaseAbsolu.Split(Path.DirectorySeparatorChar)
                Dim chemin As New System.Text.StringBuilder

                For i = 1 To frac.Count - 1
                    chemin.Append(frac(i - 1)).Append(Path.DirectorySeparatorChar)
                Next

                chemin.Remove(chemin.Length - 1, 1)

                Return chemin.ToString
            End Get
        End Property
        ''' <summary>
        ''' Nom complet absolu du résultat sélectionné
        ''' </summary>
        ''' <returns>Le chemin absolu du dossier ou du fichier sélectionné</returns>
        Public ReadOnly Property getResultatFull As String
            Get
                Return _FichierSelectionne
            End Get
        End Property
        ''' <summary>
        ''' Nom complet relatif du résultat sélectionné
        ''' </summary>
        ''' <returns>Le chemin relatif du dossier ou du fichier sélectionné</returns>
        Public ReadOnly Property getResultatRelatif As String
            Get
                Return Replace(_FichierSelectionne, _RepertoireBaseAbsolu, "..")
            End Get
        End Property
        ''' <summary>
        ''' Nom du résultat sélectionné
        ''' </summary>
        ''' <returns>Le nom du résultat sélectionné</returns>
        Public ReadOnly Property getResultatSimple As String
            Get
                Return Path.GetFileName(_FichierSelectionne)
            End Get
        End Property
        Public Property DialogBoxTexteDescription As String
            Set(ByVal value As String)
                Try
                    If IsNothing(_TextDescription) Then
                        _TextDescription = value
                    End If
                    _Form_Main_TXT_Description.Text = value
                Catch ex As Exception
                    Throw New VBToolsException(ERR_DESCRIPTION, ex)
                End Try
            End Set
            Get
                Return _Form_Main_TXT_Description.Text
            End Get
        End Property
        Public Property DialogBoxPoliceDescription As Font
            Set(ByVal value As Font)
                Try
                    If IsNothing(_TextFont) Then
                        _TextFont = value
                    End If
                    _Form_Main_TXT_Description.Font = value
                Catch ex As Exception
                    Throw New VBToolsException(ERR_DESCRIPTION, ex)
                End Try
            End Set
            Get
                Return _Form_Main_TXT_Description.Font
            End Get
        End Property
        ''' <summary>
        ''' Récupère le 1er texte inscrit dans le cadre description
        ''' </summary>
        ''' <returns>Le texte dans le cadre description d'origine</returns>
        ''' <remarks></remarks>
        Protected ReadOnly Property getDescriptionOrigineTexte As String
            Get
                Return _TextDescription
            End Get
        End Property
        ''' <summary>
        ''' Récupère la 1ere Police inscrite dans le cadre description
        ''' </summary>
        ''' <returns>La 1ere police du cadre description d'origine</returns>
        ''' <remarks></remarks>
        Protected ReadOnly Property getDescriptionOrigineFont As Font
            Get
                Return _TextFont
            End Get
        End Property
#End Region

#Region "Property Windows Forms"
        Protected Property MainForm_BTOK_Enabled As Boolean
            Get
                Return _Form_Main_BT_OK.Enabled
            End Get
            Set(ByVal value As Boolean)
                _Form_Main_BT_OK.Enabled = value
            End Set
        End Property
        Protected Property MainForm_TXTResultat_Size As Size
            Set(ByVal value As Size)
                _Form_Main_TXT_Result.Size = value
            End Set
            Get
                Return _Form_Main_TXT_Result.Size
            End Get
        End Property
        Protected Property MainForm_TXTResultat_Location As Point
            Set(ByVal value As Point)
                _Form_Main_TXT_Result.Location = value
            End Set
            Get
                Return _Form_Main_TXT_Result.Location
            End Get
        End Property
        Protected Property MainForm_TreeView_Size As Size
            Set(ByVal value As Size)
                _Form_Main_TreeView.Size = value
            End Set
            Get
                Return _Form_Main_TreeView.Size
            End Get
        End Property
        Protected Property MainForm_TreeView_Location As Point
            Set(ByVal value As Point)
                _Form_Main_TreeView.Location = value
            End Set
            Get
                Return _Form_Main_TreeView.Location
            End Get
        End Property
        Protected Property MainForm_TXTResultat_Text As String
            Get
                Return Me._Form_Main_TXT_Result.Text
            End Get
            Set(ByVal value As String)
                Me._Form_Main_TXT_Result.Text = value
            End Set
        End Property
#End Region

#Region "Constructeur"
        ''' <summary>
        ''' Ouverture d'une boite de dialogue 
        ''' </summary>
        ''' <param name="RepertoireBase">Chemin absolu du dossier ou va s'ouvrir la boite de dialogue</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal RepertoireBase As String, ByVal VoirFichier As Boolean)

            If Directory.Exists(RepertoireBase) Then
                _RepertoireBaseAbsolu = RepertoireBase
                _VoirFichierDansTreeView = VoirFichier
            Else
                Throw New VBToolsException(String.Format(ERR_DOSSIER_INCONNU, RepertoireBase))
            End If

            Try
                With _Form_Main
                    .Size = New Size(686, 508)
                    .MinimumSize = .Size
                    .MaximumSize = .Size

                    'TEXTBOX
                    Setup(_Form_Main_TXT_Result, 496, 23, 0, 440)
                    _Form_Main_TXT_Description.Multiline = True
                    Setup(_Form_Main_TXT_Description, 670, 58, 0, 0)
                    _Form_Main_TXT_Description.ReadOnly = True
                    _Form_Main_TXT_Result.ReadOnly = True
                    _TextColor = _Form_Main_TXT_Description.ForeColor

                    'TREEVIEW
                    Setup(_Form_Main_TreeView, 670, 370, 0, 66)
                    RemplissageTreeView()

                    'BOUTON
                    Setup(_Form_Main_BT_Annuler, 75, 23, 502, 440)
                    _Form_Main_BT_Annuler.Text = "Annuler"
                    Setup(_Form_Main_BT_OK, 75, 23, 583, 440)
                    _Form_Main_BT_OK.Text = "OK"
                    _Form_Main_BT_OK.Enabled = False

                    'FORM PRINCIPALE
                    .Text = _RepertoireBaseAbsolu
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_CREATE_WINFORM, ex)
            End Try
        End Sub
#End Region

#Region "Gestion ChoixFichier en Externe"
        Public Sub ShowDialog()
            Try
                _Form_Main.ShowDialog()
            Catch ex As Exception
                Throw New VBToolsException(ERR_SHOWDIALOG, ex)
            End Try
        End Sub
#End Region

#Region "Gestion des controles"
        ''' <summary>
        ''' Création d'un control dans la fenetre principale
        ''' </summary>
        ''' <param name="Ctrl">nom du control</param>
        ''' <param name="W">Size Width</param>
        ''' <param name="H">Size Height</param>
        ''' <param name="X">Location X</param>
        ''' <param name="Y">Location Y</param>
        ''' <remarks></remarks>
        Protected Sub Setup(ByVal Ctrl As Control, ByVal W As Integer, ByVal H As Integer, ByVal X As Integer, ByVal Y As Integer)
            Try
                _Form_Main.Controls.Add(Ctrl)
                With Ctrl
                    .Size = New Size(W, H)
                    .Location = New Point(X, Y)
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_CREATIONCONTROLE, ex)
            End Try
        End Sub
        Protected Sub RemplissageTreeView(Optional ByRef listeExtention As List(Of String) = Nothing)
            Try
                _Form_Main_TreeView.Nodes.Clear()
                With _Form_Main_TreeView
                    .TopNode = .Nodes.Add(_RepertoireBaseAbsolu, getRepertoireBaseRelatif)
                    'boucle sur tout les répertoires du dossier courant
                    For Each Rep As String In Directory.GetDirectories(_RepertoireBaseAbsolu)
                        'ajoute le dossier dans le treeview
                        .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                        'entre dans le dossier
                        NextNode(Rep, _Form_Main_TreeView.TopNode, listeExtention)
                    Next
                    If _VoirFichierDansTreeView Then
                        'boucle sur tout les fichiers du dossier courant
                        For Each Fichier As String In Directory.GetFiles(_RepertoireBaseAbsolu)
                            If Not Path.GetFileName(Fichier).Chars(0) = "~" Then
                                If IsNothing(listeExtention) Then
                                    .TopNode.Nodes.Add(Path.GetFileName(Fichier))
                                Else
                                    If listeExtention.Contains(Path.GetExtension(Fichier).ToLower) Then
                                        .TopNode.Nodes.Add(Path.GetFileName(Fichier))
                                    End If
                                End If
                            End If
                        Next
                    End If
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_CHARGETREEVIEW, ex)
            End Try
        End Sub
        Private Sub NextNode(ByVal Repertoire As String, ByVal NodeActuel As TreeNode, ByRef listeExtention As List(Of String))
            Try
                Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)

                For Each rep As String In Directory.GetDirectories(Repertoire)
                    Node.Nodes.Add(rep, Path.GetFileName(rep))
                    NextNode(rep, Node, listeExtention)
                Next
                If _VoirFichierDansTreeView Then
                    For Each Fichier As String In Directory.GetFiles(Repertoire)
                        If Not Path.GetFileName(Fichier).Chars(0) = "~" Then
                            If IsNothing(listeExtention) Then
                                Node.Nodes.Add(Path.GetFileName(Fichier))
                            Else
                                If listeExtention.Contains(Path.GetExtension(Fichier).ToLower) Then
                                    Node.Nodes.Add(Path.GetFileName(Fichier))
                                End If
                            End If

                        End If
                    Next
                End If
            Catch ex As Exception
                Throw New VBToolsException(ERR_CHARGETREEVIEW, ex)
            End Try
        End Sub
#End Region

#Region "Evenement"
        Private Sub BTCancel_click() Handles _Form_Main_BT_Annuler.Click
            _FichierSelectionne = Nothing
            _Form_Main.Close()
        End Sub
        Private Sub BTOK_click() Handles _Form_Main_BT_OK.Click
            _Form_Main.Close()
        End Sub

        Protected Overridable Sub TreeViewSelectionFichier(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles _Form_Main_TreeView.AfterSelect
            Dim selection As New System.Text.StringBuilder

            With selection
                .Append(getRepertoireBeforeBase)
                .Append(Path.DirectorySeparatorChar)
                .Append(_Form_Main_TreeView.SelectedNode.FullPath)
            End With

            If _VoirFichierDansTreeView Then
                If File.Exists(selection.ToString) Then
                    _FichierSelectionne = selection.ToString
                    _Form_Main_TXT_Result.Text = _Form_Main_TreeView.SelectedNode.Text
                    _Form_Main_BT_OK.Enabled = True
                Else
                    _FichierSelectionne = Nothing
                    _Form_Main_TXT_Result.Text = String.Empty
                    _Form_Main_BT_OK.Enabled = False
                End If
            Else
                _FichierSelectionne = selection.ToString
                _Form_Main_TXT_Result.Text = _Form_Main_TreeView.SelectedNode.Text
                _Form_Main_BT_OK.Enabled = True
            End If
        End Sub
#End Region
    End Class

    '-----------------------------------------------
    '   Classe : BOXSAVEFILE
    '-----------------------------------------------
    Public Class BoxSaveFile
        Inherits ChoiceBox

        Private Const ERR_NOMETHODEEXT = "La méthode de gestion de l'extention n'est pas connue"
        Private Const ERR_NOTEXT = "{0} n'est pas une extention de fichier valable"
        Private Const ERR_NOMNOTGOOD = "{0} n'est pas un nom de fichier reconnu"

        'si le fichier existe déjà
        Private Const ERR_SAMEFILE = "Le fichier {0} existe déjà!!!"
        Private _TextDescFont As New Font("Arial", 12.0, FontStyle.Italic)

        Private _nomExtention As String
        Private _nomFichierParDefault As String
        Private _selectionTreeView As String

        'Element Windows Forms
        Private WithEvents TXT_Fichier As New TextBox
        Private TXT_Ext As New TextBox

        Enum ext As Integer
            RemplaceExtention = 1
            AdditionneExtention = 2
        End Enum

#Region "Constructeur"
        ''' <summary>
        ''' Sélection d'un dossier pour enregistrement d'un fichier
        ''' </summary>
        ''' <param name="RepertoireRacine">Répertoire Racine lors de l'enregistrement</param>
        ''' <param name="NomFichierParDefaut">Fichier avec extension non modifiable</param>
        ''' <remarks></remarks>
        Sub New(ByVal RepertoireRacine As String, ByVal NomFichierParDefaut As String)
            MyClass.New(RepertoireRacine, NomFichierParDefaut, Path.GetExtension(NomFichierParDefaut), ext.RemplaceExtention)
        End Sub
        ''' <summary>
        ''' Sélection d'un dossier pour enregistrement d'un fichier
        ''' </summary>
        ''' <param name="RepertoireRacine">Répertoire Racine lors de l'enregistrement</param>
        ''' <param name="NomFichierAvecExtention">Fichier avec extension</param>
        ''' <param name="Extension">.ext</param>
        ''' <param name="action">remplace l'extension existante (.txt->.ext) ou l'ajoute (.txt->.txt.ext)</param>
        ''' <remarks></remarks>
        Sub New(ByVal RepertoireRacine As String, ByVal NomFichierAvecExtention As String, ByVal Extension As String, ByVal action As ext)
            MyBase.New(RepertoireRacine, False)

            '(note : création classe service avec check regex?)
            Dim regex As New System.Text.RegularExpressions.Regex("^\.[a-zA-Z0-9_]+")
            If Not regex.Match(Extension).Success Then
                Throw New VBToolsException(String.Format(ERR_NOTEXT, Extension))
            End If

            Dim regex2 As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9\s\-_éàè\+]+[\.a-zA-Z0-9]+")
            If Not regex2.Match(NomFichierAvecExtention).Success Then
                Throw New VBToolsException(String.Format(ERR_NOMNOTGOOD, NomFichierAvecExtention))
            End If

            If action = ext.RemplaceExtention Then
                _nomExtention = Extension
                _nomFichierParDefault = Path.GetFileNameWithoutExtension(NomFichierAvecExtention)
            ElseIf action = ext.AdditionneExtention Then
                _nomExtention = Path.GetExtension(NomFichierAvecExtention) & Extension
                _nomFichierParDefault = Path.GetFileNameWithoutExtension(NomFichierAvecExtention)
            Else
                Throw New VBToolsException(ERR_NOMETHODEEXT)
            End If

            CtrlConstruction()
        End Sub
        Private Sub CtrlConstruction()
            'TREEVIEW
            MainForm_TreeView_Size = New Size(MainForm_TreeView_Size.Width, MainForm_TreeView_Size.Height - 29)

            'TEXTBOX
            'on copie les coordonnées de TXT_Result avant de la changer de place
            Setup(TXT_Fichier, MainForm_TXTResultat_Size.Width - 46, MainForm_TXTResultat_Size.Height, MainForm_TXTResultat_Location.X, MainForm_TXTResultat_Location.Y)
            MainForm_TXTResultat_Location = New Point(MainForm_TXTResultat_Location.X, MainForm_TXTResultat_Location.Y - 29)
            MainForm_TXTResultat_Size = New Size(MainForm_TreeView_Size.Width, MainForm_TXTResultat_Size.Height)

            With TXT_Fichier
                Setup(TXT_Ext, 40, .Size.Height, .Location.X + .Size.Width + 6, .Location.Y)
                TXT_Ext.ReadOnly = True
            End With

            TXT_Ext.Text = _nomExtention
            TXT_Fichier.Text = _nomFichierParDefault

        End Sub
#End Region

#Region "Region des controls"
        Private Sub changeTextDescriptionSameFile(ByVal maDescription As String, ByVal monFont As Font, ByVal maCouleur As Color)
            MyBase.DialogBoxPoliceDescription = monFont
            MyBase.DialogBoxTexteDescription = maDescription
            Me.TXT_Fichier.BackColor = maCouleur
        End Sub
#End Region

#Region "Evènement choix utilisateur"
        Private Sub traitementChoixUtilisateur()
            Dim cheminAbs As New System.Text.StringBuilder
            Dim nomFichier As String = TXT_Fichier.Text.Trim
            'détermination du nom absolu du fichier choisi
            With cheminAbs
                .Append(getRepertoireBeforeBase)
                .Append(Path.DirectorySeparatorChar)
                .Append(_selectionTreeView)
                .Append(Path.DirectorySeparatorChar)
                .Append(nomFichier)
                .Append(_nomExtention)
            End With

            'si fichier existe
            If File.Exists(cheminAbs.ToString) Then
                changeTextDescriptionSameFile(String.Format(ERR_SAMEFILE, cheminAbs.ToString), _TextDescFont, Color.Red)
                MyBase.MainForm_BTOK_Enabled = False
            Else
                changeTextDescriptionSameFile(MyBase.getDescriptionOrigineTexte, MyBase.getDescriptionOrigineFont, Color.White)
                MyBase.setFichierSelectionne = cheminAbs.ToString
                MyBase.MainForm_BTOK_Enabled = True
            End If
        End Sub
        Protected Overrides Sub TreeViewSelectionFichier(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)
            _selectionTreeView = e.Node.FullPath
            MainForm_TXTResultat_Text = _selectionTreeView
            traitementChoixUtilisateur()
        End Sub
        Private Sub TXTFichier_TextChanged() Handles TXT_Fichier.TextChanged
            Me.TXT_Fichier.Text = Me.TXT_Fichier.Text.Trim()
            traitementChoixUtilisateur()
        End Sub
#End Region

    End Class

    '-----------------------------------------------
    '   Classe : BOXOPENFILE
    '-----------------------------------------------
    Public Class BoxOpenFile
        Inherits ChoiceBox
        Private Const _EXT_ALL_DESCRIPTION = "Tous (*.*)"
        Private Const _EXT_SPECIF_DESCRIPTION = "Type ({0})"
        Private Const _ERR_INITIALISATION = "Erreur pendant l'ouverture de l'objet BoxOpenFile"

        'élément Windows form ajoutés
        Private WithEvents CB As New ComboBox
        Private TXT_CBox As New TextBox

        Protected Friend _Close As Boolean = False
        Private _Ext As New List(Of String)
        Private _ListExtAPrendre As List(Of String) = Nothing

#Region "property"
        ''' <summary>
        ''' Liste d'extention de fichier à afficher de la forme (.ext)
        ''' </summary>
        ''' <value>doit être conforme à la regex : ^\.[a-zA-Z0-9_]+</value>
        ''' <returns>la liste des extentions</returns>
        ''' <remarks></remarks>
        Public Property listeExtention As String()
            Get
                If IsNothing(_ListExtAPrendre) Then
                    Return Nothing
                Else
                    Return _ListExtAPrendre.ToArray
                End If
            End Get
            Set(ByVal value As String())
                If Not IsNothing(value) Then
                    Dim _ListExtAPrendre As New List(Of String)
                    Dim desc As New System.Text.StringBuilder()

                    For Each ext As String In value
                        Dim regex As New System.Text.RegularExpressions.Regex("^\.[a-zA-Z0-9_]+")
                        If regex.Match(ext).Success Then
                            _ListExtAPrendre.Add(ext.ToLower)
                            desc.Append(ext.ToLower).Append("|")
                        End If
                    Next

                    'retire le dernier caractère
                    desc.Remove(desc.Length - 1, 1)

                    TXT_CBox.Text = String.Format(_EXT_SPECIF_DESCRIPTION, desc.ToString)
                Else
                    _ListExtAPrendre = Nothing
                    TXT_CBox.Text = _EXT_ALL_DESCRIPTION
                End If

                MyBase.RemplissageTreeView(_ListExtAPrendre)

            End Set
        End Property
#End Region

#Region "Constructeur"
        Sub New(ByVal RepertoireRacine As String)
            MyBase.New(RepertoireRacine, True)
            Try
                'TEXTBOX
                MainForm_TXTResultat_Size = New Size(MainForm_TXTResultat_Size.Width - 177, MainForm_TXTResultat_Size.Height)

                Setup(TXT_CBox, 171, 23, 325, 440)
                With TXT_CBox
                    .Enabled = False
                    .Text = _EXT_ALL_DESCRIPTION
                    .BringToFront()
                End With

                MyBase.RemplissageTreeView(_ListExtAPrendre)

            Catch ex As Exception
                Throw New VBToolsException(_ERR_INITIALISATION, ex)
            End Try
        End Sub
#End Region
    End Class
End Namespace