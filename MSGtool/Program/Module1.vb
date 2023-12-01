Imports SolidWorks.Interop.sldworks
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System
Imports System.Management
Imports System.Security.Cryptography
Imports System.Text
Imports System.Net.NetworkInformation

Imports System.Runtime.InteropServices

Module Module1
    Public MySwApp As SldWorks
    Public MathUtil As MathUtility
    Public Factor(0, 0), prsion As Double
    Public Mater(), SzmStr(100) As String, bRet(20) As Boolean, ie As Integer

    Public dllPath, Fpth As String
    Public PitX As Double = 2, Xo, Fio As Double, Iox As Byte
    'Dim Fpth As String = dllPath & "SZM.data"
    Public Sub ss()

    End Sub

    Public Function getSwApp() As Boolean
        If IsNothing(MySwApp) Then
            Try
                MySwApp = GetObject(, "SldWorks.Application")
            Catch ex As Exception
                Return False
            End Try
        End If

        Return True
        MySwApp.CommandInProgress = True
        '执行exe程序之前先告诉SOLIDWORKS准备在进程外将进行一系列API调用这样提高了进程外应用程序的性能
    End Function

    Public Function FacDat(Mater As String) As Boolean

        Dim sbj() As String

        Dim File As String = dllPath & "系数\" & Mater & ".Factor"
        If Not IO.File.Exists(File) Then File = dllPath & "系数\未指定.Factor"
        Try
            Using Sr As StreamReader = New StreamReader(File)
                Dim Str As String
                Dim line As String
                Do
                    line = Sr.ReadLine()
                    Str = Str & "	" & line
                Loop Until line Is Nothing
                Sr.Close()
                sbj = Split(Str, "	")
            End Using

            ReDim Factor((UBound(sbj) - 2) / 5 - 1, 4)

            Dim ia, ib As Integer
            For i = 1 To UBound(sbj) - 1
                Factor(ib, ia) = sbj(i)
                ia = ia + 1
                If ia > 4 Then ia = 0 : ib = ib + 1
            Next

        Catch ex As Exception
            MsgBox("系数文件不存在或错误")
            Return False
        End Try
        Return True
    End Function


    Public Function Regis() As Boolean
        Return True
        Dim br As BinaryReader
        Try
            Dim code As String
            Dim query As SelectQuery = New SelectQuery("select * from Win32_ComputerSystemProduct")
            br = New BinaryReader(New FileStream(dllPath & "MSG.license", FileMode.Open))
            Dim Cypher As String = br.ReadString()
            br.Close()
            Using searcher As ManagementObjectSearcher = New ManagementObjectSearcher(query)
                For Each item In searcher.Get
                    code = item("UUID").ToString()
                    If code = RSA_Decrypt(Cypher) Then
                        Return True
                    Else
                        Return False
                    End If
                Next
            End Using
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''解密
    Public Function RSA_Decrypt(Cypher As String) As String

        Dim DataToDecrypt As Byte() = Convert.FromBase64String(Cypher)

        Try
            Dim RSA As RSACryptoServiceProvider = New RSACryptoServiceProvider()
            'RSA.ImportParameters(RSAKeyInfo)
            Dim PrivateKey_bytes As Byte() = Convert.FromBase64String("BwIAAACkAABSU0EyAAQAAAEAAQD1Cg0spl2a5mKkM2Uchif/lu4jJNyO73P2zMCvLyZCRY+EtMFxb/ZJWVYrspBYDU4S16kQ9cxV+K0ITOFcfM/5AT4dmrGgVq+VIosREPUmJbPuAYf7eq5Iiy8EUI382NGUjVQYdsDZzwl2HB3hwJHlevF0J4bUVBM8ypWuHYqf4bNZjZwA3K3Ff3T9Og4gRSJUa8a4O4iCd7jvxdcChuVDauu77OFCa2+yh439s8p0Gf2SmMVbEm/kpQc1m5JOR/y3ZCjYjpsZeXJf3KyvBVPCXbQkjWQTIIdRYmk2pdAxX9EjBVNSZ/Xk7CfTkMk0uxLUpAFY5wG/cv3d+gaVkvPkn/OnlgogT/BNHLdNBS6f4kXpJWDtrn73PG4nrIYpgp8g19j3179dJVV016vJ44Z9Aje/9FoBwLIcVzCy0CmyUfe1KIgvQbrhpxQZF4lvCB3eTGqf2AZkHoTkeQt4wlEpvWaOG45iB7tTWHg29nAkxDl4mg0CBRL6ydwcu5rR7hgu1NoP0HML0XNbqECnVUu5EelrCVa1vopI14jxtNVzlfFgu3MVCc2yQLngc98Fa3cxYZdKidjy0e8wary+13kKjYCvzaZ2OLnrLvntwNm6L6Tf3+0Lbu4/120f1gyl7JqHekrjBp4LckKSWG6rZDi4apFqA9GYriarlMZqdj6eA0JO/fBRS/RvaLsZ1yUcKYmUQHDWtQZqV+Z3jzqxBV5A5aLq0zdqQXjZ7LQztpPRLEokbZIAvCzvsG5rXt0rcLY=")
            RSA.ImportCspBlob(PrivateKey_bytes)

            'OAEP padding Is only available on Microsoft Windows XP Or later. 
            Dim bytes_Plain As Byte() = RSA.Decrypt(DataToDecrypt, False)
            Dim ByteConverter As UnicodeEncoding = New UnicodeEncoding()
            Dim str_Plain As String = ByteConverter.GetString(bytes_Plain)
            Return str_Plain
        Catch ex As CryptographicException
            Return Nothing
        End Try

    End Function

#Region "二进制文件读写"

    Sub WriterData()

        Dim uni As New UnicodeEncoding()
        Dim bw As BinaryWriter

        Try '尝试实例化 ，创建二进制文件
            bw = New BinaryWriter(New FileStream(Fpth, FileMode.Create))

            For Each bl As Boolean In bRet
                bw.Write(bl)
            Next

            For Each str As String In SzmStr
                If str Is Nothing Then Exit For
                bw.Write(str)
            Next

        Catch ex As Exception
            'MsgBox(ex.Message + "\n Cannot create file.")
        End Try

        bw.Close()
    End Sub
    Sub ReaderData()

        If Not File.Exists(Fpth) Then Exit Sub

        Dim uni As New UnicodeEncoding()
        Dim br As BinaryReader

        Try
            br = New BinaryReader(New FileStream(Fpth, FileMode.Open))

            For i = 0 To UBound(bRet)
                bRet(i) = br.ReadBoolean()
            Next

            Dim ia As Integer = 0
            Do
                SzmStr(ia) = br.ReadString()
                ia = ia + 1
            Loop

        Catch ex As Exception
            'MsgBox(ex.Message + "\n Cannot create file.")
        End Try

        br.Close()
        Mater = Split(SzmStr(0), "/")
    End Sub
#End Region


#Region "测试代码"


    Sub 孤立编辑()
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 2 Then Exit Sub
        Dim AssyDoc As AssemblyDoc = Model
        Dim EditModel As ModelDoc2 = AssyDoc.GetEditTarget
        Dim Comp As Component2 = AssyDoc.GetEditTargetComponent
        Comp.Select4(False, Nothing, False)
        ' AssyDoc.EditPart2(False, False, Nothing)
        Model.Extension.RunCommand(2732, "")
        Model.Extension.RunCommand(2726, "")
        'Dim Entityf As Entity = Comp
        '16没有value = instance.Isolate()孤立接口
    End Sub


    Dim traverseLevel As Long


    Sub mainDS()

        Dim myModel As ModelDoc2 = MySwApp.ActiveDoc
        Dim featureMgr As FeatureManager = myModel.FeatureManager
        Dim rootNode As TreeControlItem = featureMgr.GetFeatureTreeRootItem2(1)
        If Not IsNothing(rootNode) Then
            Debug.Print("")
            traverseLevel = 0
            traverse_node(rootNode)
        End If
    End Sub


    Private Sub traverse_node(node As TreeControlItem)

        Dim childNode As TreeControlItem
        Dim featureNode As Feature
        Dim componentNode As Component2
        Dim restOfString As String
        Dim indent As String
        Dim i As Long
        Dim compName As String
        Dim suppr As Long, supprString As String
        Dim vis As Long, visString As String
        Dim fixed As Boolean, fixedString As String
        Dim componentDoc As Object, docString As String
        Dim refConfigName As String
        Dim displayNodeInfo As Boolean = False
        Dim nodeObject As Object = node.Object

        Dim nodeObjectType As Long = node.ObjectType

        For i = 1 To traverseLevel
            indent = indent & "  "
        Next i
        If (displayNodeInfo) Then Debug.Print(indent & node.Text & " : " & restOfString)

        traverseLevel = traverseLevel + 1
        childNode = node.GetFirstChild()
        While Not childNode Is Nothing
            traverse_node(childNode)
            childNode = childNode.GetNext
        End While
        traverseLevel = traverseLevel - 1
    End Sub


    Sub maiddn()
        Dim swApp As SldWorks
        Dim swModel As ModelDoc2
        Dim swEqnMgr As EquationMgr
        Dim i As Long

        swModel = MySwApp.ActiveDoc
        swEqnMgr = swModel.GetEquationMgr

        Dim nCount As Long = swEqnMgr.GetCount
        For i = 0 To nCount - 1
            Debug.Print("  Equation(" & i & ")  = " & swEqnMgr.Equation(i))
            Debug.Print("    Value = " & swEqnMgr.Value(i))
            Debug.Print("    Index = " & swEqnMgr.Status)
            Debug.Print("    Global variable? " & swEqnMgr.GlobalVariable(i))
        Next i
    End Sub

#End Region


End Module


Partial Class SolidWorksMacro


    Public WithEvents pDoc As PartDoc

    Public WithEvents aDoc As AssemblyDoc

    Public WithEvents dDoc As DrawingDoc


    Public Sub maincc()


        Dim swModel As ModelDoc2


        swModel = swApp.ActiveDoc


        ' Determine the document type


        ' and set up event  handlers

        If swModel.GetType = 1 Then

            pDoc = swModel

        ElseIf swModel.GetType = 2 Then

            aDoc = swModel

        ElseIf swModel.GetType = 3 Then

            dDoc = swModel

        End If


        AttachEventHandlers()


    End Sub


    Sub AttachEventHandlers()
        AttachSWEvents()
    End Sub



    Sub AttachSWEvents()
        If Not pDoc Is Nothing Then
            AddHandler pDoc.UserSelectionPostNotify, AddressOf Me.pDoc_UserSelectionPostNotify
        End If

        If Not aDoc Is Nothing Then
            AddHandler aDoc.UserSelectionPostNotify, AddressOf Me.aDoc_UserSelectionPostNotify
        End If

        If Not dDoc Is Nothing Then
            AddHandler dDoc.UserSelectionPostNotify, AddressOf Me.dDoc_UserSelectionPostNotify
        End If

    End Sub



    Private Function pDoc_UserSelectionPostNotify() As Integer
        MsgBox("Entity selected in part document.")
    End Function



    Public Function aDoc_UserSelectionPostNotify() As Integer
        MsgBox("Entity selected in assembly document.")
    End Function



    Private Function dDoc_UserSelectionPostNotify() As Integer
        MsgBox("Entity selected in drawing document.")
    End Function



    ''' <summary>
    ''' The SldWorks swApp variable is pre-assigned for you.
    ''' </summary>
    Public swApp As SldWorks

End Class



'在窗口显示钣金零件的参数
'增加对钣金零件厚度的修改
'增加智能零件库
'智能门板插件
'增加草图智能链接尺寸
'增加装配体智能方程
'文件交错浏览
'创建孔特征
'创建钣金切口特征

'Dim computerProperties As IPGlobalProperties = IPGlobalProperties.GetIPGlobalProperties()
'Dim nics As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
'Console.WriteLine("Interface information for {0}.{1}     ", computerProperties.HostName, "0000" & computerProperties.DomainName)
'Dim adapter As NetworkInterface
'For Each adapter In nics
'    Dim properties As IPInterfaceProperties = adapter.GetIPProperties()
'    Console.WriteLine(adapter.Description)
'    Console.WriteLine(String.Empty.PadLeft(adapter.Description.Length, "="c))
'    Console.WriteLine("Interface type ........: {0}", adapter.NetworkInterfaceType)
'    Console.WriteLine("Physical Address ......: {0}", adapter.GetPhysicalAddress())
'Next adapter
