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

        Private Const ERR_CREATE_WINFORM = "Erreur lors de la construction de la fenêtre de dialogue"
        Private Const ERR_DOSSIER_INCONNU = "Le chemin absolu {0} n'existe pas ou n'est pas accessible"
        Private Const ERR_CREATIONCONTROLE = "Ereur lors de la création d'un control dans la fenetre principale"
        Private Const ERR_DESCRIPTION = "Ereur lors du formatage de la zone Texte de description"
        Private Const ERR_SHOWDIALOG = "Ereur lors de l'affichage de la boite de dialogue"
        Private Const ERR_CHARGETREEVIEW = "Ereur lors de la récupération des noms de fichier pour remplir la boite de dialogue"
        Private Const ERR_LASTMETHODE = "Ereur lors de l'utilisation de la méthode Last()"

        Private _RepertoireBaseAbsolu As String

        Protected WithEvents _Form_Main As New Form
        Protected WithEvents _Form_Main_BT_OK As New Button
        Protected WithEvents _Form_Main_BT_Annuler As New Button
        Protected WithEvents _Form_Main_TreeView As New TreeView
        Protected _Form_Main_TXT_Result As New TextBox
        Protected _Form_Main_TXT_Description As New TextBox


        Private _TextDescription As String
        Private _TextFont As Font
        Private _TextColor As System.Drawing.Color
        Private _FichierSelectionne As String

#Region "OBSOLETE"
        <Obsolete("Cette méthode est obsolète, utiliser getFichierChoisi() en remplacement.")> Overridable ReadOnly Property Resultat As String
            Get
                Return _FichierSelectionne
            End Get
        End Property
#End Region

#Region "Property"
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
        Public ReadOnly Property getFichierChoisi As String
            Get
                Return _FichierSelectionne
            End Get
        End Property
        Public ReadOnly Property getFichierChoisi_NomSeul As String
            Get
                Return Path.GetFileName(_FichierSelectionne)
            End Get
        End Property
        Public ReadOnly Property getFichierChoisi_NomSeulSansExtention As String
            Get
                Return Path.GetFileNameWithoutExtension(_FichierSelectionne)
            End Get
        End Property
        Public ReadOnly Property getFichierChoisi_DepuisCheminRelatif As String
            Get
                Return Replace(_FichierSelectionne, _RepertoireBaseAbsolu, "..")
            End Get
        End Property
        Public Property DialogBoxTexteDescription As String
            Set(ByVal value As String)
                _TextDescription = value
                _Form_Main_TXT_Description.Text = _TextDescription
            End Set
            Get
                Return _Form_Main_TXT_Description.Text
            End Get
        End Property
        Public Property DialogBoxPoliceDescription As Font
            Set(ByVal value As Font)
                _TextFont = value
                _Form_Main_TXT_Description.Font = _TextFont
            End Set
            Get
                Return _Form_Main_TXT_Description.Font
            End Get
        End Property
#End Region

#Region "Constructeur"
        ''' <summary>
        ''' Ouverture d'une boite de dialogue 
        ''' </summary>
        ''' <param name="RepertoireBase">Chemin absolu du dossier ou va s'ouvrir la boite de dialogue</param>
        ''' <param name="RemplirTreeView">Remplir la boite de dialogue avec les fichiers présents dans les dossiers</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal RepertoireBase As String, Optional ByVal RemplirTreeView As Boolean = False)

            If Directory.Exists(RepertoireBase) Then
                _RepertoireBaseAbsolu = RepertoireBase
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
                    If RemplirTreeView Then RemplissageTreeView()

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
        Public Sub Description(ByVal Texte As String, Optional ByVal Police As Single = 8.25, Optional ByVal Style As FontStyle = FontStyle.Regular)
            Try
                With _Form_Main_TXT_Description
                    .Text = Texte
                    .Font = New Font(_Form_Main.Font.FontFamily.Name, Police, Style)
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_DESCRIPTION, ex)
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
        Protected Overridable Sub RemplissageTreeView()
            Try
                _Form_Main_TreeView.Nodes.Clear()
                With _Form_Main_TreeView
                    .TopNode = .Nodes.Add(_RepertoireBaseAbsolu, getRepertoireBaseRelatif)
                    For Each Rep As String In Directory.GetDirectories(_RepertoireBaseAbsolu)
                        .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                        NextNode(Rep, _Form_Main_TreeView.TopNode)
                    Next
                    For Each Fichier As String In Directory.GetFiles(_RepertoireBaseAbsolu)
                        If Not Last(Fichier, "\").Chars(0) = "~" Then .TopNode.Nodes.Add(Path.GetFileName(Fichier))
                    Next
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_CHARGETREEVIEW, ex)
            End Try
        End Sub
        Protected Overridable Sub NextNode(ByVal Repertoire As String, ByVal NodeActuel As TreeNode)
            Try
                Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)

                For Each rep As String In Directory.GetDirectories(Repertoire)
                    Node.Nodes.Add(rep, Path.GetFileName(rep))
                    NextNode(rep, Node)
                Next
                For Each Fichier As String In Directory.GetFiles(Repertoire)
                    If Not Last(Fichier, "\").Chars(0) = "~" Then Node.Nodes.Add(Path.GetFileName(Fichier))
                Next
            Catch ex As Exception
                Throw New VBToolsException(ERR_CHARGETREEVIEW, ex)
            End Try

        End Sub
#End Region

#Region "Gestion du Traitement chemin / répertoire / fichier"
        Protected Function Last(ByRef Mot As String, ByVal Car As String) As String
            Try
                Return Mot.Split(Car)(Mot.Split(Car).Count - 1)
            Catch ex As Exception
                Throw New VBToolsException(ERR_LASTMETHODE, ex)
            End Try
        End Function
        Protected Function AllNotLast(ByRef Mot As String, ByVal car As String) As String
            Dim texte As String = Nothing
            For i = 0 To Mot.Split(car).Count - 2
                texte += Mot.Split(car)(i)
                If Not i = Mot.Split(car).Count - 2 Then texte += car
            Next
            Return texte
        End Function
#End Region

#Region "Evenement"
        Protected Overridable Sub BTCancel_click() Handles _Form_Main_BT_Annuler.Click
            _Form_Main.Close()
        End Sub
        Protected Overridable Sub BTOK_click() Handles _Form_Main_BT_OK.Click
            _FichierSelectionne = AllNotLast(_RepertoireBaseAbsolu, "\") & "\" & _Form_Main_TreeView.SelectedNode.FullPath
            _Form_Main.Close()
        End Sub
        Protected Overridable Sub TV_AfterSelect() Handles _Form_Main_TreeView.AfterSelect
            _Form_Main_TXT_Result.Text = _Form_Main_TreeView.SelectedNode.Text
            _Form_Main_BT_OK.Enabled = True
        End Sub
#End Region

        Protected Sub ReinitialiseTextDescription()
            With Me._Form_Main_TXT_Description
                .Text = _TextDescription
                .ForeColor = _TextColor
                .Font = _TextFont
            End With
        End Sub

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
        Private _Resultat As String
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
                Return AllNotLast(getRepertoireBaseAbsolu, "\") & "\" & _Form_Main_TreeView.SelectedNode.FullPath
            End Get
        End Property
        Private ReadOnly Property NomComplet As String
            Get
                Dim NomFichier As New System.Text.StringBuilder
                With NomFichier
                    .Append(_Form_Main_TXT_Result.Text).Append(Path.DirectorySeparatorChar)
                    .Append(TXT_Fichier.Text.Trim).Append(TXT_Ext.Text)
                End With
                Return NomFichier.ToString
            End Get
        End Property
        Overrides ReadOnly Property Resultat As String
            Get
                Return _Resultat
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
            MyBase.New(RepertoireRacine)
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
            MyBase.New(RepertoireRacine)
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
            With _Form_Main_TreeView.Size
                _Form_Main_TreeView.Size = New Size(.Width, .Height - 29)
            End With

            'TEXTBOX
            With _Form_Main_TXT_Result
                'on copie les coordonnées de TXT_Result avant de la changer de place
                Setup(TXT_Fichier, .Size.Width - 46, .Size.Height, .Location.X, .Location.Y)
                _Form_Main_TXT_Result.Location = New Point(.Location.X, .Location.Y - 29)
                _Form_Main_TXT_Result.Size = New Size(_Form_Main_TreeView.Size.Width, _Form_Main_TXT_Result.Size.Height)
            End With
            With TXT_Fichier
                Setup(TXT_Ext, 40, .Size.Height, .Location.X + .Size.Width + 6, .Location.Y)
                TXT_Ext.ReadOnly = True
            End With
            RemplissageTreeView()

        End Sub
#End Region

#Region "Region des controls"
        Private Sub VerifCanClickOK()
            If Not _CompleteLoad Then Exit Sub
            _Form_Main_BT_OK.Enabled = False
            If Not TXT_Fichier.Text = vbNullString And Not _Form_Main_TXT_Result.Text = vbNullString And VerifNomFichier() Then
                If BlackListOK() Then _Form_Main_BT_OK.Enabled = True
            End If
        End Sub
        Private Function VerifNomFichier() As Boolean
            If File.Exists(NomComplet) Then
                changeTextDescriptionSameFile(NomComplet)
                Return False
            ElseIf File.Exists(NomComplet & ".crp") Then
                changeTextDescriptionSameFile(NomComplet & ".crp")
                Return False
            Else
                ReinitialiseTextDescription()
                Me.TXT_Fichier.BackColor = Color.White
                Return True
            End If
        End Function
        Private Sub changeTextDescriptionSameFile(ByVal nomFichier As String)
            With MyBase._Form_Main_TXT_Description
                .Text = String.Format(ERR_SAMEFILE, nomFichier)
                .Font = _TextDescFont
            End With
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
        Protected Overrides Sub RemplissageTreeView()
            _Form_Main_TreeView.Nodes.Clear()
            With _Form_Main_TreeView
                .TopNode = .Nodes.Add(getRepertoireBaseAbsolu, getRepertoireBaseRelatif)
                For Each Rep As String In Directory.GetDirectories(getRepertoireBaseAbsolu)
                    .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                    NextNode(Rep, _Form_Main_TreeView.TopNode)
                Next
            End With
        End Sub
        Protected Overrides Sub NextNode(ByVal Repertoire As String, ByVal NodeActuel As TreeNode)
            Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)
            For Each rep As String In Directory.GetDirectories(Repertoire)
                Node.Nodes.Add(rep, Path.GetFileName(rep))
                NextNode(rep, Node)
            Next
        End Sub
        Private Function ComparatifEXT(ByVal GetExt As String, ByVal SplitExt As String) As Boolean
            Return (Path.GetExtension(GetExt).ToLower Like ("." & SplitExt.Split(".")(SplitExt.Split(".").Count - 1))).ToString.ToLower
        End Function
#End Region

#Region "Evènement"
        Protected Overrides Sub BTOK_Click() Handles _Form_Main_BT_OK.Click
            _Resultat = DossierComplet & "\" & TXT_Fichier.Text & TXT_Ext.Text
            _Form_Main.Close()
        End Sub
        Protected Overrides Sub BTCancel_Click() Handles _Form_Main_BT_Annuler.Click
            _Form_Main.Close()
        End Sub
        Protected Overrides Sub TV_AfterSelect() Handles _Form_Main_TreeView.AfterSelect
            _Form_Main_TXT_Result.Text = DossierComplet
            VerifCanClickOK()
        End Sub
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

        Protected Friend _Close As Boolean = False
        Private WithEvents CB As New ComboBox
        Private TXT_CBox As New TextBox
        Private _FileName As New List(Of String)
        Private _Ext As New List(Of String)
        Private _ListExtAPrendre As New List(Of String)

        Public Enum Donne As Integer
            Extension = 0
            FichierSeul = 1
            Fichier = 2
            NomFichierTreeview = 3
            NomFichierComplet = 4
        End Enum

        Overrides ReadOnly Property Resultat As String
            Get
                If _FileName.Count = 0 Then
                    Return Nothing
                Else
                    Return _FileName(4)
                End If
            End Get
        End Property

#Region "Constructeur"
        Sub New(ByVal RepertoireRacine As String)
            MyBase.New(RepertoireRacine)

            _ListExtAPrendre.Add("*.*")

            'TEXTBOX
            _Form_Main_TXT_Result.Size = New Size(_Form_Main_TXT_Result.Width - 177, _Form_Main_TXT_Result.Height)

            'COMBOBOX
            Setup(CB, 171, 23, 325, 440)
            CB.DropDownStyle = ComboBoxStyle.DropDownList
            Setup(TXT_CBox, 0, 0, 0, 0)
            With TXT_CBox
                .Enabled = False
                .Location = CB.Location
                .Size = CB.Size
                .Text = "Tous (*.*)"
                .BringToFront()
            End With

            RemplissageTreeView()

        End Sub
#End Region

#Region "Region des controls"
        Protected Overrides Sub RemplissageTreeView()
            _Form_Main_TreeView.Nodes.Clear()
            With _Form_Main_TreeView
                .TopNode = .Nodes.Add(getRepertoireBaseAbsolu, getRepertoireBaseRelatif)
                For Each Rep As String In Directory.GetDirectories(getRepertoireBaseAbsolu)
                    .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                    NextNode(Rep, _Form_Main_TreeView.TopNode)
                Next
                For Each Fichier As String In Directory.GetFiles(getRepertoireBaseAbsolu)
                    If Not Last(Fichier, "\").Chars(0) = "~" Then
                        For Each EXT As String In _ListExtAPrendre
                            If ComparatifEXT(Fichier, EXT) Then
                                .TopNode.Nodes.Add(Path.GetFileName(Fichier))
                            End If
                        Next
                    End If
                Next
            End With
        End Sub
        Protected Overrides Sub NextNode(ByVal Repertoire As String, ByVal NodeActuel As TreeNode)
            Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)

            For Each rep As String In Directory.GetDirectories(Repertoire)
                Node.Nodes.Add(rep, Path.GetFileName(rep))
                NextNode(rep, Node)
            Next
            For Each Fichier As String In Directory.GetFiles(Repertoire)
                If Not Last(Fichier, "\").Chars(0) = "~" Then
                    For Each EXT As String In _ListExtAPrendre
                        If ComparatifEXT(Fichier, EXT) Then
                            Node.Nodes.Add(Path.GetFileName(Fichier))
                        End If
                    Next
                End If
            Next
        End Sub
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

            Call RemplissageTreeView()
        End Sub
#End Region

#Region "Retour Apres fermeture de la boite de dialogue"
        Public Function Result(ByVal What As Donne) As String
            If _FileName.Count = 0 Then Return Nothing
            Return _FileName(What)
        End Function
#End Region

#Region "Evenement"
        Protected Overrides Sub BTOK_Click() Handles _Form_Main_BT_OK.Click
            _Close = True
            _Form_Main.Close()
        End Sub
        Protected Overrides Sub BTCancel_Click() Handles _Form_Main_BT_Annuler.Click
            _FileName.Clear()
            _Close = True
            _Form_Main.Close()
        End Sub
        Protected Overrides Sub TV_AfterSelect() Handles _Form_Main_TreeView.AfterSelect
            Selection()
            If File.Exists(_FileName(4)) Then
                _Form_Main_TXT_Result.Text = _FileName(2)
                _Form_Main_BT_OK.Enabled = True
            Else
                _Form_Main_TXT_Result.Text = ""
                _Form_Main_BT_OK.Enabled = False
            End If
        End Sub
        Private Sub CB_SelectedIndexChanged() Handles CB.SelectedIndexChanged
            TriExt(CB.Items.IndexOf(CB.SelectedItem))
        End Sub
        Private Sub FormClose(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles _Form_Main.FormClosing
            If Not _Close Then e.Cancel = True
        End Sub
        Private Sub Selection()
            _FileName.Clear()
            With _Form_Main_TreeView.SelectedNode
                _FileName.Add(Last(.Text, ".")) 'nom extension
                _FileName.Add(AllNotLast(.Text, ".")) 'nom fichier sans ext
                _FileName.Add(.Text) 'nom fichier
                _FileName.Add(.FullPath) 'nom fichier dans treeview
                _FileName.Add(AllNotLast(getRepertoireBaseAbsolu, "\") & "\" & .FullPath) 'nom fichier dans system
            End With
        End Sub
#End Region

    End Class
End Namespace