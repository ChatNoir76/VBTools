Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing

Namespace DialogBox
    Public Class ChoiceBox

        Private Const ERR_CB_WINFROM = "Erreur lors de la construction de la fenêtre de dialogue"

        Protected WithEvents _Main As New Form
        Protected WithEvents BT_OK As New Button
        Protected WithEvents BT_Annuler As New Button
        Protected WithEvents TV As New TreeView
        Protected TXT_Result As New TextBox
        Protected TXT_Description As New TextBox
        Protected _RepAll As String
        Protected _RepBase As String
        Private _Result As String

#Region "Property"
        Overridable ReadOnly Property Resultat As String
            Get
                Return _Result
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
            _RepAll = RepertoireBase
            _RepBase = Last(_RepAll, "\")

            Try
                With _Main
                    .Size = New Size(686, 508)
                    .MinimumSize = .Size
                    .MaximumSize = .Size

                    'TEXTBOX
                    Setup(TXT_Result, 496, 23, 0, 440)
                    TXT_Description.Multiline = True
                    Setup(TXT_Description, 670, 58, 0, 0)
                    TXT_Description.ReadOnly = True
                    TXT_Result.ReadOnly = True

                    'TREEVIEW
                    Setup(TV, 670, 370, 0, 66)
                    If RemplirTreeView Then RemplissageTreeView()

                    'BOUTON
                    Setup(BT_Annuler, 75, 23, 502, 440)
                    BT_Annuler.Text = "Annuler"
                    Setup(BT_OK, 75, 23, 583, 440)
                    BT_OK.Text = "OK"
                    BT_OK.Enabled = False

                    'FORM PRINCIPALE
                    .Text = _RepAll
                End With
            Catch ex As Exception
                Throw New VBToolsException(ERR_CB_WINFROM, ex)
            End Try
        End Sub
#End Region

#Region "Gestion ChoixFichier en Externe"
        Public Sub ShowDialog()
            _Main.ShowDialog()
        End Sub
        Public Sub Description(ByVal Texte As String, Optional ByVal Police As Single = 8.25, Optional ByVal Style As FontStyle = FontStyle.Regular)
            With TXT_Description
                .Text = Texte
                .Font = New Font(_Main.Font.FontFamily.Name, Police, Style)
            End With
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
            _Main.Controls.Add(Ctrl)
            With Ctrl
                .Size = New Size(W, H)
                .Location = New Point(X, Y)
            End With
        End Sub
        Protected Overridable Sub RemplissageTreeView()
            TV.Nodes.Clear()
            With TV
                .TopNode = .Nodes.Add(_RepAll, _RepBase)
                For Each Rep As String In Directory.GetDirectories(_RepAll)
                    .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                    NextNode(Rep, TV.TopNode)
                Next
                For Each Fichier As String In Directory.GetFiles(_RepAll)
                    If Not Last(Fichier, "\").Chars(0) = "~" Then .TopNode.Nodes.Add(Path.GetFileName(Fichier))
                Next
            End With
        End Sub
        Protected Overridable Sub NextNode(ByVal Repertoire As String, ByVal NodeActuel As TreeNode)
            Dim Node As TreeNode = NodeActuel.Nodes(Repertoire)

            For Each rep As String In Directory.GetDirectories(Repertoire)
                Node.Nodes.Add(rep, Path.GetFileName(rep))
                NextNode(rep, Node)
            Next
            For Each Fichier As String In Directory.GetFiles(Repertoire)
                If Not Last(Fichier, "\").Chars(0) = "~" Then Node.Nodes.Add(Path.GetFileName(Fichier))
            Next
        End Sub
#End Region

#Region "Gestion du Traitement chemin / répertoire / fichier"
        Protected Function Last(ByRef Mot As String, ByVal Car As String) As String
            Return Mot.Split(Car)(Mot.Split(Car).Count - 1)
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
        Protected Overridable Sub BTCancel_click() Handles BT_Annuler.Click
            _Main.Close()
        End Sub
        Protected Overridable Sub BTOK_click() Handles BT_OK.Click
            _Result = AllNotLast(_RepAll, "\") & "\" & TV.SelectedNode.FullPath
            _Main.Close()
        End Sub
        Protected Overridable Sub TV_AfterSelect() Handles TV.AfterSelect
            TXT_Result.Text = TV.SelectedNode.Text
            BT_OK.Enabled = True
        End Sub
#End Region

    End Class
    Public Class BoxSaveFile
        Inherits ChoiceBox

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
                Return AllNotLast(_RepAll, "\") & "\" & TV.SelectedNode.FullPath
            End Get
        End Property
        Private ReadOnly Property NomComplet As String
            Get
                Dim NomFichier As New System.Text.StringBuilder
                With NomFichier
                    .Append(TXT_Result.Text).Append(Path.PathSeparator)
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
            With TV.Size
                TV.Size = New Size(.Width, .Height - 29)
            End With

            'TEXTBOX
            With TXT_Result
                'on copie les coordonnées de TXT_Result avant de la changer de place
                Setup(TXT_Fichier, .Size.Width - 46, .Size.Height, .Location.X, .Location.Y)
                TXT_Result.Location = New Point(.Location.X, .Location.Y - 29)
                TXT_Result.Size = New Size(TV.Size.Width, TXT_Result.Size.Height)
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
            BT_OK.Enabled = False
            If Not TXT_Fichier.Text = vbNullString And Not TXT_Result.Text = vbNullString And VerifNomFichier() Then
                If BlackListOK() Then BT_OK.Enabled = True
            End If
        End Sub
        Private Function VerifNomFichier() As Boolean

            If File.Exists(NomComplet) Then Return False
            If File.Exists(NomComplet & ".crp") Then Return False
            Return True

        End Function
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
            TV.Nodes.Clear()
            With TV
                .TopNode = .Nodes.Add(_RepAll, _RepBase)
                For Each Rep As String In Directory.GetDirectories(_RepAll)
                    .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                    NextNode(Rep, TV.TopNode)
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
        Protected Overrides Sub BTOK_Click() Handles BT_OK.Click
            _Resultat = DossierComplet & "\" & TXT_Fichier.Text & TXT_Ext.Text
            MyBase._Main.Close()
        End Sub
        Protected Overrides Sub BTCancel_Click() Handles BT_Annuler.Click
            MyBase._Main.Close()
        End Sub
        Protected Overrides Sub TV_AfterSelect() Handles TV.AfterSelect
            TXT_Result.Text = DossierComplet
            VerifCanClickOK()
        End Sub
        Protected Overridable Sub TXTFichier_TextChanged() Handles TXT_Fichier.TextChanged
            VerifCanClickOK()
        End Sub
#End Region

    End Class
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
            TXT_Result.Size = New Size(TXT_Result.Width - 177, TXT_Result.Height)

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
            TV.Nodes.Clear()
            With TV
                .TopNode = .Nodes.Add(_RepAll, _RepBase)
                For Each Rep As String In Directory.GetDirectories(_RepAll)
                    .TopNode.Nodes.Add(Rep, Path.GetFileName(Rep))
                    NextNode(Rep, TV.TopNode)
                Next
                For Each Fichier As String In Directory.GetFiles(_RepAll)
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
        Protected Overrides Sub BTOK_Click() Handles BT_OK.Click
            _Close = True
            _Main.Close()
        End Sub
        Protected Overrides Sub BTCancel_Click() Handles BT_Annuler.Click
            _FileName.Clear()
            _Close = True
            _Main.Close()
        End Sub
        Protected Overrides Sub TV_AfterSelect() Handles TV.AfterSelect
            Selection()
            If File.Exists(_FileName(4)) Then
                TXT_Result.Text = _FileName(2)
                BT_OK.Enabled = True
            Else
                TXT_Result.Text = ""
                BT_OK.Enabled = False
            End If
        End Sub
        Private Sub CB_SelectedIndexChanged() Handles CB.SelectedIndexChanged
            TriExt(CB.Items.IndexOf(CB.SelectedItem))
        End Sub
        Private Sub FormClose(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles _Main.FormClosing
            If Not _Close Then e.Cancel = True
        End Sub
        Private Sub Selection()
            _FileName.Clear()
            With TV.SelectedNode
                _FileName.Add(Last(.Text, ".")) 'nom extension
                _FileName.Add(AllNotLast(.Text, ".")) 'nom fichier sans ext
                _FileName.Add(.Text) 'nom fichier
                _FileName.Add(.FullPath) 'nom fichier dans treeview
                _FileName.Add(AllNotLast(_RepAll, "\") & "\" & .FullPath) 'nom fichier dans system
            End With
        End Sub
#End Region

    End Class
End Namespace