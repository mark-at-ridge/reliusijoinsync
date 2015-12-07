Public Class Rfc4180Writer
    Public Shared Sub WriteDataTable(ByVal sourceTable As DataTable, ByVal writer As IO.TextWriter, ByVal includeHeaders As Boolean)
        If (includeHeaders) Then
            Dim headerValues As List(Of String) = New List(Of String)()
            For Each column As DataColumn In sourceTable.Columns
                headerValues.Add(QuoteValue(column.ColumnName))
            Next
        End If

        Dim items() As String = Nothing
        For Each row As DataRow In sourceTable.Rows
            items = row.ItemArray.Select(Function(obj) QuoteValue(obj.ToString())).ToArray()
            writer.WriteLine(String.Join(",", items))
        Next

        writer.Flush()
    End Sub

    Private Shared Function QuoteValue(ByVal value As String) As String
        Return String.Concat("""", value.Replace("""", """"""), """")
    End Function

End Class
