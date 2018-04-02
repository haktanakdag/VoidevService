Imports System
Imports System.Data
Imports System.Configuration
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
Imports System.Data.OleDb
Imports System.Collections
Imports System.Timers
Imports System.IO.Compression
Imports System.Xml
Imports System.Web
Imports System.Net
Imports System.Net.Mail



Public Class DbUp
    Public Property ConfigurationManager As Object

    Public Function gGetConnectionString() As String
        Dim conn As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("connstring")
        Dim connString As String = conn.ConnectionString
        Return connString
    End Function
    Public Function queryExec(ByVal query As String) As Boolean
        Dim sonuc As Boolean = False
        Using connection As New SqlConnection(gGetConnectionString())
            connection.Open()
            Dim cmd As New SqlCommand
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.Connection = connection
            If (cmd.ExecuteNonQuery()) Then
                sonuc = True
            End If
        End Using
        Return sonuc
    End Function
    Public Function dbUpdate() As Boolean
        Dim query As String = "IF  NOT EXISTS (SELECT * FROM sys.objects  WHERE object_id = OBJECT_ID(N'[dbo].[TBLSCRIPTMERKEZ]') AND type in (N'U'))" & vbCrLf &
    "BEGIN" & vbCrLf &
    "CREATE TABLE [dbo].[TBLSCRIPTMERKEZ](" & vbCrLf &
    "[ID] [int] IDENTITY(1,1) NOT NULL," & vbCrLf &
    "[TXTSCRIPTTANIM] [varchar](100) NULL, " & vbCrLf &
    "[SUBEKOD] [int] NULL, " & vbCrLf &
    " [TXTSCRIPT] [text] NOT NULL, " & vbCrLf &
    " [DURUM] [tinyint] NULL," & vbCrLf &
    "CONSTRAINT [PK_TBLSCRIPT] PRIMARY KEY CLUSTERED " & vbCrLf &
    " ( " & vbCrLf &
    "[ID] ASC " & vbCrLf &
    ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] " & vbCrLf &
    " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] " & vbCrLf &
    "SET ANSI_PADDING OFF" & vbCrLf &
    "END"
        queryExec(query)

        query = "If EXISTS(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SSP_HA_ADISYON_SAYI]') AND type in (N'P', N'PC')) " & vbCrLf &
    "DROP PROCEDURE [dbo].[SSP_HA_ADISYON_SAYI]"

        queryExec(query)

        query = "CREATE PROC [dbo].[SSP_HA_ADISYON_SAYI] " & vbCrLf &
    "AS" & vbCrLf &
    "Select " & vbCrLf &
    "SUBE_KODU" & vbCrLf &
    ",count(AU.adisyonno) As sayi" & vbCrLf &
    ",convert(varchar(20),Case when convert(time,K_ZAMAN) <='06:00:00.0000000'  then convert(date,DATEADD(dd,-1,K_ZAMAN)) else convert(date,K_ZAMAN) end,112) tarih" & vbCrLf &
    ",SUM(yekun) BorcToplam" & vbCrLf &
    ",sum(CASE WHEN ao.odeme=1 then AO.miktar ELSE 0 END) NakitOdeme" & vbCrLf &
    ",sum(CASE WHEN ao.odeme=23 then AO.miktar else 0 END) KKOdeme" & vbCrLf &
    "From TBL_ADISYONUST AU" & vbCrLf &
    "INNER Join TBL_ODEME AO ON AO.adisyonno=AU.adisyonno" & vbCrLf &
    "WHERE Convert(Date, AU.tarih) >='20160501' " & vbCrLf &
    "GROUP BY  SUBE_KODU" & vbCrLf &
    ",case when convert(time,K_ZAMAN) <='06:00:00.0000000'  then convert(date,DATEADD(dd,-1,K_ZAMAN)) else convert(date,K_ZAMAN) end"
        queryExec(query)
        Return True
    End Function
End Class
