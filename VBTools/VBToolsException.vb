Public Class VBToolsException
    Inherits Exception

    Private Const _ERR_NOSOURCE = "Pas d'erreur source"
    Private _errSource As String

#Region "Property"
    Public ReadOnly Property getErrSource As String
        Get
            Return _errSource
        End Get
    End Property
#End Region

#Region "Constructeur"
    Public Sub New(ByVal errmsg As String, Optional ByVal errExcep As Exception = Nothing)
        MyBase.New(errmsg)
        _errSource = If(IsNothing(errExcep), _ERR_NOSOURCE, errExcep.Message)
    End Sub
#End Region



End Class
