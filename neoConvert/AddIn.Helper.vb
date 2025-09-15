Imports EPDM.Interop.epdm
Imports System.Text

Partial Public Class AddIn

    Public Function GetextensionNames(ByRef Vault As IEdmVault10) As String()

        Dim extensionNamesStoredRaw As String = String.Empty

        Dim storage As IEdmDictionary5 = Vault.GetDictionary(ADDIN_NAME, True)

        storage.StringGetAt(AddIn.STORAGEKEY, extensionNamesStoredRaw)

        Dim extensionNames As New List(Of String)

        If Not String.IsNullOrWhiteSpace(extensionNamesStoredRaw) Then

            extensionNames.AddRange(extensionNamesStoredRaw.Split(";").Select(Function(x) x.Trim()).Where(Function(x) Not String.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
        End If

        Return extensionNames.ToArray()
    End Function


End Class
