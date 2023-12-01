
Imports System.IO
Imports System.Management
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms

Public Class lic
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim code As String
            Dim query As SelectQuery = New SelectQuery("select * from Win32_ComputerSystemProduct")
            Using searcher As ManagementObjectSearcher = New ManagementObjectSearcher(query)
                For Each item In searcher.Get
                    code = item("UUID").ToString()
                    TextBox1.Text = RSA_Encrypt(code)
                Next
            End Using

        Catch ex As Exception

        End Try
    End Sub
    ''加密
    Public Function RSA_Encrypt(Text As String) As String
        Dim ByteConverter As UnicodeEncoding = New UnicodeEncoding()
        Dim DataToEncrypt As Byte() = ByteConverter.GetBytes(Text)
        Try
            Dim RSA As RSACryptoServiceProvider = New RSACryptoServiceProvider()
            'RSA.ImportParameters(RSAKeyInfo)
            Dim PublicKey_bytes As Byte() = Convert.FromBase64String("BgIAAACkAABSU0ExAAQAAAEAAQDNwjm/OMJGQSVgqvUuiZwmt1iPdTVQcG4/QFEkoSimCl39v35LOZGl+DoxlY4kHc6wIuiKd+H7CA8hDCnm+zHbjPndXm+2f7JV1hNDDNKaWrQllG5hF/zvlMEq79pq+8npDI4kRylkrPoMi/8w54S3pcf7vx4hItEDs9oZdSLuuQ==")
            RSA.ImportCspBlob(PublicKey_bytes)
            'OAEP padding Is only available on Microsoft Windows XP Or later. 
            Dim Cypher_bytes As Byte() = RSA.Encrypt(DataToEncrypt, False)
            Dim Cypher_str As String = Convert.ToBase64String(Cypher_bytes)
            Return Cypher_str

        Catch ex As CryptographicException
            Return Nothing
        End Try
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim bw As BinaryWriter
        Dim license As String = TextBox2.Text
        If license = "" Then MsgBox("请输入注册码") : Exit Sub

        Try '尝试实例化 ，创建二进制文件

            bw = New BinaryWriter(New FileStream(dllPath & "MSG.license", FileMode.Create))
            bw.Write(license)
            bw.Close()
        Catch ex As Exception

        End Try

        If Not Regis() Then TextBox2.Text = Nothing : MsgBox("注册码错误")

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = Nothing
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'TextBox1.SelectAll()
        Clipboard.SetDataObject(TextBox1.Text)
    End Sub

End Class