'mettre le fichier texte à jour pour le changement de version
<Assembly: Reflection.AssemblyFileVersion("1.0.1.0")> 
<Assembly: Reflection.AssemblyVersion("1.0.1.0")> 
<Assembly: Reflection.AssemblyInformationalVersion("Build 2")> 

Module _VBTools
    'application console pour test de la dll
    'changer le type d'application pour la génération
    Sub main()
        Console.WriteLine("---Début DU PROGRAMME---")
        Console.Read()
        Try
            Dim maBox As New DialogBox.BoxSaveFile("N:\API\DOMAINE DE TRAVAIL R&D\Personnel\Julien\Interface ModeOp2", "Archivage.pdf", ".pdf", DialogBox.BoxSaveFile.ext.Remplace)
            maBox.ShowDialog()
            Console.WriteLine(maBox.Resultat)
            Console.WriteLine(maBox.getFichierChoisi_NomSeul)
            Console.WriteLine(maBox.getFichierChoisi_DepuisCheminRelatif)

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
End Module


