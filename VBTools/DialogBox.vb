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
        'Private Const ERR_LASTMETHODE = "Ereur lors de l'utilisation de la méthode Last()"

        'extention autorisé à l'affichage
        Private Const _EXT_ALL = "*.*"

        Private _RepertoireBaseAbsolu As String
        Private _FichierSelectionne As String
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
        Protected WriteOnly Property RepertoireBaseAbsolu As String
            Set(ByVal value As String)
                _RepertoireBaseAbsolu = value
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
        Protected Sub BTCancel_click() Handles _Form_Main_BT_Annuler.Click
            _FichierSelectionne = Nothing
            _Form_Main.Close()
        End Sub
        Protected Sub BTOK_click() Handles _Form_Main_BT_OK.Click
            _Form_Main.Close()
        End Sub
        Protected Sub TreeViewSelectionFichier() Handles _Form_Main_TreeView.AfterSelect
            Dim selection As New System.Text.StringBuilder
            With selection
                .Append(_RepertoireBaseAbsolu)
                .Append(Path.DirectorySeparatorChar)
                .Append(_Form_Main_TreeView.SelectedNode.FullPath)
            End With
            _FichierSelectionne = selection.ToString
            _Form_Main_TXT_Result.Text = _Form_Main_TreeView.SelectedNode.Text
            _Form_Main_BT_OK.Enabled = True
        End Sub
#End Region
    End Class

    '-----------------------------------------------
    '   Classe : BOXSAVEFILE
    '-----------------------------------------------
    Public Class BoxSaveFile
        Inherits ChoiceBox

        Private Const ERR_SAMEFILE = "Le fichier {0} existe déjà!!!"
        Private _TextDescFont As New Font("Arial", 12.0, FontStyle.Italic)
        Private _TextDescRed As Color = Color.Red

        Private WithEvents TXT_Fichier As New TextBox
        Private TXT_Ext As New TextBox
        Private _CompleteLoad As Boolean = False
        Private _listeNomInterdit As List(Of String)

        Enum ext As Integer
            Remplace = 1
            Ajoute = 2
            Autre = 3
        End Enum

#Region "Property"
        ''' <summary>
        ''' Liste de nom de fichier interdit lors de l'enregistrement (nom complet avec extention)
        ''' </summary>
        ''' <value>Liste nom de fichier interdit</value>
        ''' <returns>Liste nom de fichier interdit</returns>
        ''' <remarks></remarks>
        Public Property ListeNomInterdit As List(Of String)
            Get
                Return _listeNomInterdit
            End Get
            Set(ByVal value As List(Of String))
                _listeNomInterdit = value
            End Set
        End Property
        Private ReadOnly Property DossierComplet As String
            Get
                Return Nothing 'AllNotLast(getRepertoireBaseAbsolu, "\") & "\" & _Form_Main_TreeView.SelectedNode.FullPath
            End Get
        End Property
        Private ReadOnly Property NomComplet As String
            Get
                Dim NomFichier As New System.Text.StringBuilder
                With NomFichier
                    '.Append(_Form_Main_TXT_Result.Text).Append(Path.DirectorySeparatorChar)
                    .Append(TXT_Fichier.Text.Trim).Append(TXT_Ext.Text)
                End With
                Return NomFichier.ToString
            End Get
        End Property
#End Region

#Region "Constructeur"
        ''' <summary>
        ''' Sélection d'un dossier pour enregistrement d'un fichier
        ''' </summary>
        ''' <param name="RepertoireRacine">Répertoire Racine lors de l'enregistrement</param>
        ''' <param name="NomFichierParDefaut">Fichier avec extension non modifiable</param>
        ''' <remarks></remarks>
        Sub New(ByVal RepertoireRacine As String, Optional ByVal NomFichierParDefaut As String = Nothing)
            MyBase.New(RepertoireRacine, False)
            CtrlConstruction()

            If Not IsNothing(NomFichierParDefaut) Then
                TXT_Ext.Text = Path.GetExtension(NomFichierParDefaut)
                TXT_Fichier.Text = Path.GetFileNameWithoutExtension(NomFichierParDefaut)
            Else
                TXT_Ext.Visible = False
                TXT_Fichier.Size = New Size(TXT_Fichier.Size.Width + TXT_Ext.Size.Width + 6, TXT_Fichier.Size.Height)
            End If

            _CompleteLoad = True
        End Sub
        ''' <summary>
        ''' Sélection d'un dossier pour enregistrement d'un fichier
        ''' </summary>
        ''' <param name="RepertoireRacine">Répertoire Racine lors de l'enregistrement</param>
        ''' <param name="NomFichierParDefaut">Fichier avec extension</param>
        ''' <param name="Extension">.ext</param>
        ''' <param name="action">remplace l'extension existante (.txt->.ext) ou l'ajoute (.txt->.txt.ext)</param>
        ''' <remarks></remarks>
        Sub New(ByVal RepertoireRacine As String, ByVal NomFichierParDefaut As String, ByVal Extension As String, ByVal action As ext)
            MyBase.New(RepertoireRacine, False)
            CtrlConstruction()
            Select Case action
                Case 1
                    TXT_Ext.Text = Extension
                    TXT_Fichier.Text = Path.GetFileNameWithoutExtension(NomFichierParDefaut)
                Case 2
                    TXT_Ext.Text = Path.GetExtension(NomFichierParDefaut) & Extension
                    TXT_Fichier.Text = Path.GetFileNameWithoutExtension(NomFichierParDefaut)
                Case Else
                    TXT_Ext.Text = Extension
                    TXT_Fichier.Text = NomFichierParDefaut
            End Select

            _CompleteLoad = True
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
                'RemplissageTreeView()

        End Sub
#End Region

#Region "Region des controls"
        Private Sub VerifCanClickOK()
            If Not _CompleteLoad Then Exit Sub
            MainForm_BTOK_Enabled = False
            'If Not TXT_Fichier.Text = vbNullString And Not _Form_Main_TXT_Result.Text = vbNullString And VerifNomFichier() Then
            If BlackListOK() Then MainForm_BTOK_Enabled = True
            'End If
        End Sub
        Private Function VerifNomFichier() As Boolean
            If File.Exists(NomComplet) Then
                changeTextDescriptionSameFile(NomComplet)
                Return False
            ElseIf File.Exists(NomComplet & ".crp") Then
                changeTextDescriptionSameFile(NomComplet & ".crp")
                Return False
            Else
                'ReinitialiseTextDescription()
                Me.TXT_Fichier.BackColor = Color.White
                Return True
            End If
        End Function
        Private Sub changeTextDescriptionSameFile(ByVal nomFichier As String)
            'With MyBase._Form_Main_TXT_Description
            '.Text = String.Format(ERR_SAMEFILE, nomFichier)
            '.Font = _TextDescFont
            'End With
            Me.TXT_Fichier.BackColor = _TextDescRed
        End Sub
        Private Function BlackListOK() As Boolean
            'si la liste est null ou vide, pas de comparaison
            If IsNothing(_listeNomInterdit) Then Return True
            If _listeNomInterdit.Count = 0 Then Return True

            'comparaison
            For Each Nom As String In _listeNomInterdit
                If DossierComplet & "\" & Nom = NomComplet Then Return False
            Next

            Return True
        End Function
        Protected Sub NextNodeOLD(ByVal Repertoire As String, ByVal NodeActuel As TreeNode)
            Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)
            For Each rep As String In Directory.GetDirectories(Repertoire)
                Node.Nodes.Add(rep, Path.GetFileName(rep))
                NextNodeOLD(rep, Node)
            Next
        End Sub
        Private Function ComparatifEXT(ByVal GetExt As String, ByVal SplitExt As String) As Boolean
            Return (Path.GetExtension(GetExt).ToLower Like ("." & SplitExt.Split(".")(SplitExt.Split(".").Count - 1))).ToString.ToLower
        End Function
#End Region

#Region "Evènement"
        'Protected Sub TreeViewSelectionFichierOLD() Handles _Form_Main_TreeView.AfterSelect
        '_Form_Main_TXT_Result.Text = DossierComplet
        'VerifCanClickOK()
        'End Sub
        Protected Overridable Sub TXTFichier_TextChanged() Handles TXT_Fichier.TextChanged
            VerifCanClickOK()
        End Sub
#End Region

    End Class

    '-----------------------------------------------
    '   Classe : BOXOPENFILE
    '-----------------------------------------------
    Public Class BoxOpenFile
        Inherits ChoiceBox
        Private Const _EXT_ALL_DESCRIPTION = "Tous (*.*)"
        Private Const _ERR_INITIALISATION = "Erreur pendant l'ouverture de l'objet BoxOpenFile"

        'élément Windows form ajoutés
        Private WithEvents CB As New ComboBox
        Private TXT_CBox As New TextBox

        Protected Friend _Close As Boolean = False
        Private _Ext As New List(Of String)
        Private _ListExtAPrendre As New List(Of String)

#Region "Constructeur"
        Sub New(ByVal RepertoireRacine As String)
            MyBase.New(RepertoireRacine, True)
            Try
                'ajout des extentions par défaut
                _ListExtAPrendre.Add(".pdf")

                'TEXTBOX
                MainForm_TXTResultat_Size = New Size(MainForm_TXTResultat_Size.Width - 177, MainForm_TXTResultat_Size.Height)

                'COMBOBOX
                Setup(CB, 171, 23, 325, 440)
                CB.DropDownStyle = ComboBoxStyle.DropDownList
                Setup(TXT_CBox, 0, 0, 0, 0)
                With TXT_CBox
                    .Enabled = False
                    .Location = CB.Location
                    .Size = CB.Size
                    .Text = _EXT_ALL_DESCRIPTION
                    .BringToFront()
                End With

                MyBase.RemplissageTreeView(_ListExtAPrendre)

            Catch ex As Exception
                Throw New VBToolsException(_ERR_INITIALISATION, ex)
            End Try
        End Sub
#End Region

#Region "Region des controls"
        Private Function ComparatifEXT(ByVal GetExt As String, ByVal SplitExt As String) As Boolean
            Return (Path.GetExtension(GetExt).ToLower Like ("." & SplitExt.Split(".")(SplitExt.Split(".").Count - 1))).ToString.ToLower
        End Function
#End Region

#Region "Choix Externe"
        ''' <summary>
        ''' Choix des extensions des fichiers recherchés
        ''' </summary>
        ''' <param name="Ext">exemple : Office Documents |*.doc;*.docx;*.dotx;*.dotm</param>
        ''' <remarks></remarks>
        Public Sub ChoixExtension(ByVal ParamArray Ext() As String)
            Dim nb As Integer
            For Each element As String In Ext
                nb += 1
                CB.Items.Add(element.Split("|")(0))
                _Ext.Add(element.Split("|")(1))
            Next
            If nb > 0 Then TriExt(0)
        End Sub
        Private Sub TriExt(ByVal Index As Integer) 'index correspondant à celui de la combobox et de la liste _EXT
            'on efface la liste
            _ListExtAPrendre.Clear()

            TXT_CBox.Text = CB.Items(Index)

            For Each element In _Ext(Index).Split(";")
                _ListExtAPrendre.Add(element)
            Next

            'Call RemplissageTreeView()
        End Sub
#End Region

#Region "Evenement"
        Private Sub CB_SelectedIndexChanged() Handles CB.SelectedIndexChanged
            TriExt(CB.Items.IndexOf(CB.SelectedItem))
        End Sub
        Private Sub Selection()
            'MyBase.resultat
            '_FileName.Clear()
            'With _Form_Main_TreeView.SelectedNode
            '_FileName.Add(Last(.Text, ".")) 'nom extension
            '_FileName.Add(AllNotLast(.Text, ".")) 'nom fichier sans ext
            '_FileName.Add(.Text) 'nom fichier
            '_FileName.Add(.FullPath) 'nom fichier dans treeview
            '_FileName.Add(AllNotLast(getRepertoireBaseAbsolu, "\") & "\" & .FullPath) 'nom fichier dans system
            'End With
        End Sub
#End Region

    End Class
End Namespace