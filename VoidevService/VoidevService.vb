Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Linq
Imports System.ServiceProcess
Imports System.Text
Imports System.Threading
Imports System.IO
Imports System.Configuration
Imports System.Data.OleDb
Imports System.Collections

Imports System.Timers
Imports System.IO.Compression
Imports System.Xml
Imports System.Web
Imports System.Net
Imports System.Net.Mail
Public Class VoidevService
    Dim writer As StreamWriter
    Dim durum As Boolean
    Public Property ConfigurationManager As Object

    Protected Overrides Sub OnStart(ByVal args() As String)
        'logyaz("Servis Başladı")
        Timer1.Start()
    End Sub

    Protected Overrides Sub OnStop()
        'logyaz("Servis Bitti")
    End Sub

    Public Function gGetConnectionString() As String
        Dim conn As ConnectionStringSettings = CType(ConfigurationManager.ConnectionStrings("connstring"), ConnectionStringSettings)
        Dim connString As String = conn.ConnectionString
        Return connString
    End Function
    Public Function mailGonder()
        Dim mail As New MailMessage()
        Dim SmtpServer As New SmtpClient
        SmtpServer.Credentials = New Net.NetworkCredential("helpdesk@aybel.com.tr", "helpdesk123")
        SmtpServer.Port = 587
        SmtpServer.EnableSsl = True
        SmtpServer.Host = "smtp.yandex.com.tr"
        mail.From = New MailAddress("helpdesk@aybel.com.tr")
        mail.To.Add("haktan.akdag@gmail.com")
        mail.Subject = "Konu Deneme"
        mail.Body = "Mesaj deneme"
        SmtpServer.Send(mail)
    End Function
    Public Function logyaz(ByVal yazilacaklog As String) As Boolean
        Dim strParam As Dictionary(Of String, String)
        strParam = GetXmlParam()
        Dim fs As FileStream = New FileStream(strParam.Item("logdosyaad").ToString(), FileMode.OpenOrCreate)
        writer = New StreamWriter(fs)
        writer.WriteLine(yazilacaklog + System.DateTime.Now.ToString())
        writer.Flush()
        writer.Close()
    End Function

    Private Sub Islem()
        logyaz("işlemlere girdi")
        'mailGonder()
        'Dim dbup As New DbUp()
        'DbUp.dbUpdate()
        'WSSelectScriptler()
        'ExecScriptler()
        'WSInsertTutarlilik()
    End Sub

    Private Sub Timer1_Elapsed(sender As Object, e As ElapsedEventArgs) Handles Timer1.Elapsed
        logyaz("Timer devreye girdi")
        Dim strParam As Dictionary(Of String, String)
        strParam = GetXmlParam()
        Dim ts As String
        Dim tm As String
        ts = System.DateTime.Now.Hour.ToString()
        If (System.DateTime.Now.Hour.ToString().Length = 1) Then
            ts = "0" + System.DateTime.Now.Hour.ToString()
        End If
        tm = System.DateTime.Now.Minute.ToString()
        If (System.DateTime.Now.Minute.ToString().Length = 1) Then
            tm = "0" + System.DateTime.Now.Minute.ToString()
        End If
        Dim saat As String = ts + ":" + tm
        If (strParam.Item("saat").ToString() = saat) Then
            Islem()
            If (strParam.Item("yedekleme").ToString() = True) Then
                Dim yedeklenendosya As String = Yedekle(strParam.Item("vtad").ToString(), strParam.Item("yedekdosyayol").ToString())
                Dim dosya As String() = yedeklenendosya.Split(New Char() {"\"c})
                UploadFile(yedeklenendosya + ".gz", strParam.Item("ftpadres").ToString() + dosya(3) + ".gz", strParam.Item("ftpkulad").ToString(), strParam.Item("ftpsifre").ToString())
            End If
        End If
    End Sub
    Private Function Yedekle(ByVal dataadi As String, ByVal datayolu As String) As String
        logyaz("Yedekleme başladı")
        KlasorYoksaOlustur(datayolu)
        Dim yedekdosya As String = "'" + datayolu + dataadi + System.DateTime.Now.ToString().Replace(" ", "").Replace(".", "").Replace(":", "") + ".BAK'"
        Using connection As New SqlConnection(gGetConnectionString())
            connection.Open()
            Dim cmd As New SqlCommand
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "BACKUP DATABASE " + dataadi + " TO DISK=" + yedekdosya
            cmd.Connection = connection
            If (cmd.ExecuteNonQuery()) Then
                logyaz("Yedekleme bitti")
            End If
        End Using
        Dim directorySelected As New DirectoryInfo(datayolu)
        logyaz("Sıkıştırma Başladı")
        Compress(directorySelected, datayolu)
        logyaz("Sıkıştırma Bitti")
        logyaz("Bak Dosyası Siliniyor")
        DosyaSil(datayolu, yedekdosya)
        logyaz("Bak Dosyası Silindi")
        Dim yedeklenendosya As String = yedekdosya.Replace("'", "")
        Return yedeklenendosya
    End Function
    Private Function DosyaSil(ByVal dir, ByVal yedekdosya) As Boolean
        Dim fileToRemove As String
        Try
            yedekdosya = yedekdosya.ToString().Replace("'", "")
            Dim dosya As String() = yedekdosya.Split(New Char() {"\"c})
            Dim paths() As String = IO.Directory.GetFiles(dir, dosya(3).ToString())
            For i As Integer = 0 To paths.Length
                fileToRemove = paths(i).ToString
                System.IO.File.Delete(fileToRemove)
                If (Not System.IO.File.Exists(fileToRemove)) Then
                    Return True
                Else
                    writer.WriteLine("Dosya Silinemedi.")
                    writer.Flush()
                    Return False
                End If
            Next
            Return False
        Catch ex As Exception
            writer.WriteLine(ex.ToString())
            writer.Flush()
            Return False
        End Try
        Return False
    End Function
    Public Sub Compress(directorySelected As DirectoryInfo, directoryPath As String)
        For Each fileToCompress As FileInfo In directorySelected.GetFiles()
            Using originalFileStream As FileStream = fileToCompress.OpenRead()
                If (File.GetAttributes(fileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden And fileToCompress.Extension <> ".gz" Then
                    Using compressedFileStream As FileStream = File.Create(fileToCompress.FullName & ".gz")
                        Using compressionStream As New GZipStream(compressedFileStream, CompressionMode.Compress)
                            originalFileStream.CopyTo(compressionStream)

                        End Using
                    End Using
                End If
            End Using
        Next
    End Sub

    Private Sub KlasorYoksaOlustur(datayolu As String)
        Dim directoryName As String = Path.GetDirectoryName(datayolu)
        If (directoryName.Length > 0) AndAlso (Not Directory.Exists(directoryName)) Then
            Directory.CreateDirectory(directoryName)
        End If
    End Sub
    Public Sub UploadFile(ByVal _FileName As String, ByVal _UploadPath As String, ByVal _FTPUser As String, ByVal _FTPPass As String)
        Dim _FileInfo As New System.IO.FileInfo(_FileName)
        Dim _FtpWebRequest As System.Net.FtpWebRequest = CType(System.Net.FtpWebRequest.Create(New Uri(_UploadPath)), System.Net.FtpWebRequest)

        _FtpWebRequest.Credentials = New System.Net.NetworkCredential(_FTPUser, _FTPPass)
        _FtpWebRequest.KeepAlive = False

        _FtpWebRequest.Timeout = 20000
        _FtpWebRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile
        _FtpWebRequest.UseBinary = True
        _FtpWebRequest.ContentLength = _FileInfo.Length

        ' The buffer size is set to 2kb
        Dim buffLength As Integer = 2048
        Dim buff(buffLength - 1) As Byte
        Dim _FileStream As System.IO.FileStream = _FileInfo.OpenRead()

        Try
            Dim _Stream As System.IO.Stream = _FtpWebRequest.GetRequestStream()
            Dim contentLen As Integer = _FileStream.Read(buff, 0, buffLength)
            Do While contentLen <> 0
                _Stream.Write(buff, 0, contentLen)
                contentLen = _FileStream.Read(buff, 0, buffLength)
            Loop
            _Stream.Close()
            _Stream.Dispose()
            _FileStream.Close()
            _FileStream.Dispose()
        Catch ex As Exception
            writer.WriteLine(ex.ToString())
            writer.Flush()
        End Try
    End Sub
    Public Function GetXmlParam() As Dictionary(Of String, String)
        Dim doc As New XmlDocument
        doc.Load("D:\Parametreler.xml")
        Dim ReturnValue As New Dictionary(Of String, String)
        For Each node As XmlNode In doc.SelectNodes("//Parameters")
            ReturnValue.Add("vtad", node.SelectSingleNode("vtad").InnerText)
            ReturnValue.Add("yedekleme", node.SelectSingleNode("yedekleme").InnerText)
            ReturnValue.Add("yedekdosya", node.SelectSingleNode("yedekdosya").InnerText)
            ReturnValue.Add("loglama", node.SelectSingleNode("loglama").InnerText)
            ReturnValue.Add("logdosya", node.SelectSingleNode("logdosya").InnerText)
            ReturnValue.Add("ftpadres", node.SelectSingleNode("ftpadres").InnerText)
            ReturnValue.Add("ftpkulad", node.SelectSingleNode("ftpkulad").InnerText)
            ReturnValue.Add("ftpsifre", node.SelectSingleNode("ftpsifre").InnerText)
            ReturnValue.Add("dbupdate", node.SelectSingleNode("dbupdate").InnerText)
            ReturnValue.Add("haberlesme", node.SelectSingleNode("haberlesme").InnerText)
            ReturnValue.Add("webserviceadres", node.SelectSingleNode("webserviceadres").InnerText)
            ReturnValue.Add("saat", node.SelectSingleNode("saat").InnerText)
            ReturnValue.Add("subekod", node.SelectSingleNode("subekod").InnerText)
        Next
        Return ReturnValue
    End Function
    Public Function WSInsertTutarlilik() As Boolean
        Dim s As New WebReference.WebService()
        Dim strParam As Dictionary(Of String, String)
        strParam = GetXmlParam()
        s.Url = strParam.Item("webserviceadres").ToString() + "?op=Insert"
        Using connection As New SqlConnection(gGetConnectionString())
            connection.Open()
            Dim cmd As SqlCommand = New SqlCommand("SSP_HA_ADISYON_SAYI", connection)
            cmd.CommandType = CommandType.StoredProcedure
            Using r = cmd.ExecuteReader()
                If r.Read() Then
                    Dim i As Int32 = 0
                    While r.Read()
                        s.Insert(r(0).ToString, r(2).ToString, r(1).ToString, r(3).ToString)
                    End While
                End If
            End Using
        End Using
        Return True
    End Function
    Public Function WSSelectScriptler() As Boolean
        Dim s As New WebReference.WebService()
        Dim strParam As Dictionary(Of String, String)
        strParam = GetXmlParam()
        s.Url = strParam.Item("webserviceadres").ToString() + "?op=SelectScript"
        Dim dtList As DataTable = s.SelectScript(strParam.Item("subekod").ToString())

        Dim drlist As New List(Of DataRow)()

        For Each row As DataRow In dtList.Rows
            drlist.Add(CType(row, DataRow))
        Next row

        For Each row As DataRow In dtList.Rows
            For Each column As DataColumn In dtList.Columns
                drlist.Add(row)
            Next
        Next

        Dim Connection As New SqlConnection(gGetConnectionString())
        Connection.Open()
        Dim cmd As New SqlCommand
        cmd.Connection = Connection

        For i As Int32 = 1 To dtList.Rows.Count() Step 1
            Dim txtscript As String = drlist(i)("TXTSCRIPT").ToString().Replace("'", "''")
            cmd.CommandText = "INSERT INTO TBLSCRIPTSUBE(TXTSCRIPTTANIM, SUBEKOD,TXTSCRIPT,DURUM) VALUES('" + drlist(i)("TXTSCRIPTTANIM").ToString() + "','" + drlist(i)("SUBEKOD").ToString() + "','" + txtscript + "'," + drlist(i)("DURUM").ToString() + ")"
            'MessageBox.Show(drlist(i)("ID").ToString())
            'MessageBox.Show(drlist(i)("TXTSCRIPTTANIM").ToString())
            'MessageBox.Show(drlist(i)("SUBEKOD").ToString())
            'MessageBox.Show(drlist(i)("TXTSCRIPT").ToString())
            'MessageBox.Show(drlist(i)("DURUM").ToString())

            cmd.ExecuteNonQuery()
        Next
        Connection.Close()
        Return True
    End Function
    Public Function ExecScriptler() As Boolean
        Dim Connection As New SqlConnection(gGetConnectionString())
        Connection.Open()
        Dim cmd As New SqlCommand("SELECT ID,TXTSCRIPT,TXTSCRIPTTANIM FROM TBLSCRIPTSUBE WHERE DURUM =0 ORDER BY ID ASC")

        Dim sda As New SqlDataAdapter()
        cmd.Connection = Connection
        sda.SelectCommand = cmd
        Dim dt As New DataTable()
        dt.TableName = "TBLSCRIPTSUBE"
        sda.Fill(dt)
        Dim dr As New List(Of DataRow)()

        For Each row As DataRow In dt.Rows
            dr.Add(CType(row, DataRow))
        Next row

        For Each row As DataRow In dt.Rows
            For Each column As DataColumn In dt.Columns
                dr.Add(row)
            Next
        Next
        Dim calismadurum As Int32 = 0
        Dim s As New WebReference.WebService()
        For i As Int32 = 1 To dt.Rows.Count() Step 1
            cmd.CommandText = dr(i)("TXTSCRIPT").ToString()
            Try
                cmd.ExecuteNonQuery()
                'Script Çalışıyor Mu bakıyoruz. Eğer çalışmazsa exception kısmına düşüyor ve çalışmadı bilgisini update ediyoruz.
                cmd.CommandText = "UPDATE TBLSCRIPTSUBE SET DURUM = 1 WHERE ID = " + dr(i)("ID").ToString()
                cmd.ExecuteNonQuery()
                s.UpdateScript(dr(i)("TXTSCRIPTTANIM").ToString(), 1)
            Catch ex As Exception
                cmd.CommandText = "UPDATE TBLSCRIPTSUBE SET DURUM = 2 WHERE ID = " + dr(i)("ID").ToString()
                cmd.ExecuteNonQuery()
                s.UpdateScript(dr(i)("TXTSCRIPTTANIM").ToString(), 2)
            End Try

        Next
        Connection.Close()
        Return True
    End Function
End Class
