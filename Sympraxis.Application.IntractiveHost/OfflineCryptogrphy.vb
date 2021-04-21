Imports System.IO
Imports System.Security.Cryptography

Public Class OfflineCryptogrphy
    'Protected Sub EncryptFile(sender As Object, e As EventArgs)
    '    'Get the Input File Name and Extension.
    '    Dim fileName As String = System.Windows.Shapes.Path.GetFileNameWithoutExtension(FileUpload1.PostedFile.FileName)
    '    Dim fileExtension As String = System.Windows.Shapes.Path.GetExtension(FileUpload1.PostedFile.FileName)

    '    'Build the File Path for the original (input) and the encrypted (output) file.
    '    Dim input As String = Convert.ToString(Server.MapPath("~/Files/") & fileName) & fileExtension
    '    Dim output As String = Convert.ToString((Server.MapPath("~/Files/") & fileName) + "_enc") & fileExtension

    '    'Save the Input File, Encrypt it and save the encrypted file in output path.
    '    FileUpload1.SaveAs(input)
    '    Me.Encrypt(input, output)

    '    'Download the Encrypted File.
    '    Response.ContentType = FileUpload1.PostedFile.ContentType
    '    Response.Clear()
    '    Response.AppendHeader("Content-Disposition", "attachment; filename=" + System.Windows.Shapes.Path.GetFileName(output))
    '    Response.WriteFile(output)
    '    Response.Flush()

    '    'Delete the original (input) and the encrypted (output) file.
    '    File.Delete(input)
    '    File.Delete(output)

    '    Response.End()
    'End Sub

    'Protected Sub DecryptFile(sender As Object, e As EventArgs)
    '    'Get the Input File Name and Extension
    '    Dim fileName As String = System.Windows.Shapes.Path.GetFileNameWithoutExtension(FileUpload1.PostedFile.FileName)
    '    Dim fileExtension As String = System.Windows.Shapes.Path.GetExtension(FileUpload1.PostedFile.FileName)

    '    'Build the File Path for the original (input) and the decrypted (output) file
    '    Dim input As String = Convert.ToString(Server.MapPath("~/Files/") & fileName) & fileExtension
    '    Dim output As String = Convert.ToString((Server.MapPath("~/Files/") & fileName) + "_dec") & fileExtension

    '    'Save the Input File, Decrypt it and save the decrypted file in output path.
    '    FileUpload1.SaveAs(input)
    '    Me.Decrypt(input, output)

    '    'Download the Decrypted File.
    '    Response.Clear()
    '    Response.ContentType = FileUpload1.PostedFile.ContentType
    '    Response.AppendHeader("Content-Disposition", "attachment; filename=" + System.Windows.Shapes.Path.GetFileName(output))
    '    Response.WriteFile(output)
    '    Response.Flush()

    '    'Delete the original (input) and the decrypted (output) file.
    '    File.Delete(input)
    '    File.Delete(output)

    '    Response.End()
    'End Sub

    Private Shared Function Assign(Of T)(ByRef source As T, ByVal value As T) As T
        source = value
        Return value
    End Function

    Public Sub Encrypt(inputFilePath As String, outputfilePath As String)
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using fs As New FileStream(outputfilePath, FileMode.Create)
                Using cs As New CryptoStream(fs, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    Using fsInput As New FileStream(inputFilePath, FileMode.Open)
                        Dim data As Integer
                        While (Assign(data, fsInput.ReadByte())) <> -1
                            cs.WriteByte(CByte(data))
                        End While
                    End Using
                End Using
            End Using
        End Using
    End Sub

    Public Sub Decrypt(inputFilePath As String, outputfilePath As String)
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using fs As New FileStream(inputFilePath, FileMode.Open)
                Using cs As New CryptoStream(fs, encryptor.CreateDecryptor(), CryptoStreamMode.Read)
                    Using fsOutput As New FileStream(outputfilePath, FileMode.Create)
                        Dim data As Integer
                        While (Assign(data, cs.ReadByte())) <> -1
                            fsOutput.WriteByte(CByte(data))
                        End While
                    End Using
                End Using
            End Using
        End Using
    End Sub
End Class
