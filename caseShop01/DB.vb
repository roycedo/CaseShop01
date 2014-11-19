Public Class DB
    Public Shared ReadOnly Property connectString() As String
        Get
            Return "Server=localhost;Port=3306;Database=test;Uid=root;Pwd=roycemysql123;"
        End Get
    End Property
End Class
