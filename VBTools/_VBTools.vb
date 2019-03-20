'mettre le fichier texte à jour pour le changement de version
<Assembly: Reflection.AssemblyVersion("1.1.0.0")> 

Module _VBTools
    'application console pour test de la dll
    'changer le type d'application pour la génération
    Sub main()
        Console.WriteLine("---Début DU PROGRAMME---")
        Console.Read()
        Try
            Dim maBox As New DialogBox.ChoiceBox("D:\interfaceModeOp2")
            maBox.ShowDialog()
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
        End Try
        Console.WriteLine("---FIN DU PROGRAMME---")
        Console.Read()
    End Sub
End Module


