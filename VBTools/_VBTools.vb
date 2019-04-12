Imports System.Windows.Forms
Imports VBTools.GestionDataGridView

'mettre le fichier texte à jour pour le changement de version
<Assembly: Reflection.AssemblyFileVersion("1.2.0.2")> 
<Assembly: Reflection.AssemblyVersion("1.2.0.2")> 
<Assembly: Reflection.AssemblyInformationalVersion("Build 8")> 

Module _VBTools
    'application console pour test de la dll
    'changer le type d'application pour la génération
    Sub main()
        Console.WriteLine("---Début DU PROGRAMME---")
        Console.Read()

        Try

            testDialogBox()

            Console.Read()
        Catch ex As VBToolsException
            Console.WriteLine("VBToolsException")
            Console.WriteLine(ex.Message)
            Console.WriteLine("---Source---")
            Console.WriteLine(ex.getErrSource)
            Console.Read()
        Catch ex As Exception
            Console.WriteLine("Erreur générale")
            Console.WriteLine(ex.Message)
            Console.Read()
        Finally
            Console.WriteLine("---FIN DU PROGRAMME---")
            Console.Read()
        End Try

    End Sub

    Sub testDialogBox()
        'Dim maBox As New DialogBox.BoxSaveFile("N:\API\DOMAINE DE TRAVAIL R&D\Personnel\Julien\Interface ModeOp2", "MCD.pdf") With {.DialogBoxTexteDescription = "Ceci est le texte à définir"}
        Dim maBox As New DialogBox.BoxOpenFile("N:\API\DOMAINE DE TRAVAIL R&D\Personnel\Julien\Interface ModeOp2") With {.listeExtention = {".pdf"}}
        'Dim maBox As New DialogBox.ChoiceBox("N:\API\DOMAINE DE TRAVAIL R&D\Personnel\Julien\Interface ModeOp2", False)

        maBox.ShowDialog()
        Console.WriteLine(maBox.getResultatFull)
        Console.WriteLine(maBox.getResultatRelatif)
        Console.WriteLine(maBox.getResultatSimple)
        Console.WriteLine(maBox.getRepertoireBaseAbsolu)
        Console.WriteLine(maBox.getRepertoireBaseRelatif)

    End Sub

    Sub testReflexionClasseAnonyme()
        Dim _form As New Form
        Dim maListe As New List(Of Object)
        Dim monDGV As New DataGridView()
        _form.Controls.Add(monDGV)

        For i = 1 To 20
            maListe.Add(New With {.id = i,
                                 .name = "nom " & i,
                                 .surname_un_truc_assez_gros_pour_foutre_le_bordel = "abcdefghi0abcdefghi0abcdefghi0abcdefghi" & i})
        Next

        monDGV.DataSource = maListe

        Dim monImp As New PrintDataGridView(monDGV)

        monImp.textePremierePage = "Ceci est un texte de présentation de la liste"

        monImp.Impression()

    End Sub

    Sub testDGV()
        Dim _form As New Form
        Dim maTable As New DataTable("Ma Data Table")
        Dim monDGV As New DataGridView()

        For i = 1 To 5
            maTable.Columns.Add("colonne " & i, GetType(String))
        Next

        For l = 1 To 30
            maTable.Rows.Add("ligne " & l, "c2l" & l, "3", "4", "5")
        Next

        _form.Controls.Add(monDGV)
        monDGV.DataSource = maTable

        Dim monImp As New PrintDataGridView(monDGV)
        monImp.Impression()
    End Sub
End Module



