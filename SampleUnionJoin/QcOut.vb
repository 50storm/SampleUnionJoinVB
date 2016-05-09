Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml.Linq


Public Class QcOut

#Region "メンバー"
    Private NotCommonTables As New ArrayList
    Private SqlCommon As String

    Private _CommonDataSetName As String
    Property CommonDataSetName As String
        Get
            Return _CommonDataSetName

        End Get
        Set(value As String)
            _CommonDataSetName = value

        End Set
    End Property

    Private _UniqueDataSetName As String
    Property UniqueDataSetName As String
        Get
            Return _UniqueDataSetName
        End Get
        Set(value As String)
            _UniqueDataSetName = value
        End Set
    End Property

#End Region

#Region "コンストラクタ"

    Sub New()

        ''TODO
        ''外だしのファイルから取得するようにする
        'NotCommonTables.Add("qc_out_ucs_doa")
        'NotCommonTables.Add("qc_out_2004_noncqap")
        'NotCommonTables.Add("qc_out_2004_2900isr")
        'NotCommonTables.Add("qc_out_2004_cius")

        Try



            'XMLに共通でないテーブルが書いてあるので首都kう
            Dim Tables As XElement = XElement.Load(My.Settings("NotCmnTablesPath").ToString().Trim())

            'XMLをLinQにセット。
            Dim q = (From x In Tables.<Table>).ToArray()
            For Each table As String In q
                '共通でないテーブル名をセット・
                NotCommonTables.Add(table)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            MessageBox.Show("NotCmnTables.xmlのパスを確認してください。")

            Throw
        End Try

    End Sub

#End Region

#Region "プライベートメソッド"

    ''' <summary>
    ''' 共通のテーブル名を取得するSQLを返す
    ''' </summary>
    ''' <param name="NotCommonTable">共通でないテーブル名</param>
    ''' <returns>共通のテーブル名を取得するSQLを返す</returns>
    Private Function GetSqlCommonTableName(NotCommonTable As ArrayList) As String

        Dim Sql As String = "SELECT name FROM sysobjects WHERE name like 'qc_out_%' AND name  <> 'qc_out_2004_ucs_backup' "

        If NotCommonTable.Count <> 0 Then
            For i As Integer = 0 To NotCommonTable.Count - 1
                Sql &= " AND name <> '" & NotCommonTable.Item(i).ToString() & "' "
            Next
        End If


        Return Sql

    End Function


    ''' <summary>
    ''' テーブル名(共通）を取得する
    ''' </summary>
    ''' <param name="SqlCommonTableName"></param>
    ''' <returns></returns>
    Private Function GetCommonTableName(SqlCommonTableName As String) As ArrayList
        Dim CommonTableNames As New ArrayList

        Dim Adapter As New SqlDataAdapter
        Dim DataSetObj As New DataSet

        Try
            Using connection As New SqlConnection(My.Settings.Item("connectionString").ToString())
                connection.Open()
                Dim command As New SqlCommand()
                command.Connection = connection
                command.CommandText = SqlCommonTableName
                Adapter.SelectCommand = command


                Adapter.Fill(DataSetObj, "TableName")
                If (DataSetObj.Tables.Count = 0) Then
                    Throw New Exception("テーブル名が取得できませんでした")
                Else
                    'ArrayListにテーブル名を入れる
                    For Each Table As DataTable In DataSetObj.Tables
                        For Each Row As DataRow In Table.Rows
                            'Console.WriteLine(Row.Item(0))
                            For Each Col As DataColumn In Table.Columns
                                'Console.WriteLine(Row(Col))
                                CommonTableNames.Add(Row(Col).ToString())
                            Next
                        Next
                    Next
                End If

            End Using
        Catch ex As SqlException
            MessageBox.Show(ex.Message)
            Throw

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Throw

        End Try

        Return CommonTableNames
    End Function


#End Region

#Region "公開メソッド"

#Region "Getter"
#Region "共通のテーブル"

    ''Java Likeな書き方ではなく.NET的なプロパティーを使うように変更（わかりやすい？かもだから）
    '''' <summary>
    '''' 共通で使えるテーブルで仕様してるデータセット名をリターン
    '''' </summary>
    '''' <returns></returns>
    'Public Function geCmnDataSetName() As String
    '    Return CommonDataSetName
    'End Function

    'Public Function getUniqueDataSetName() As String
    '    Return UniqueDataSetName
    'End Function

#End Region

#Region "単一のテーブル"

#End Region
#End Region

#Region "共通のテーブル(UNION)できるもの"


    ''' <summary>
    ''' 共通テーブル（UNION）のデータセットを返す
    ''' </summary>
    ''' <param name="_DataSetName">データセットで使う名前</param>
    ''' <param name="ParaTables"></param>
    ''' <returns></returns>
    Public Function GetCommonTables(_DataSetName As String, ParamArray ByVal ParaTables() As String) As DataSet
        Dim ReturnVale As New DataSet

        Dim Sql As String = ""

        Try

            If _DataSetName.Trim() = String.Empty Then
                Throw New Exception("データセットで使う名前をセットしていください。")
            Else
                CommonDataSetName = _DataSetName
            End If


            'UNIONで共通でもってこれるテーブルを取得
            Dim TableNameCommon As ArrayList = GetCommonTableName(GetSqlCommonTableName(NotCommonTables))
            If (TableNameCommon.Count = 0) Then
                Throw New Exception("UNION結合で使うテーブルが1つも取得できなかった")
            End If


            '入力されたテーブル名がUNION可能かチェック
            Dim IsContained As New ArrayList

            For i As Integer = 0 To ParaTables.Length - 1
                If TableNameCommon.Contains(ParaTables(i)) Then
                    IsContained.Add(True)
                Else
                    IsContained.Add(False)

                End If
            Next
            If IsContained.Contains(False) Then
                Throw New Exception("UNIONで使えるテーブルではありません。")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Throw
        End Try


        Dim Adapter As New SqlDataAdapter
        Try
            Using connection As New SqlConnection(My.Settings.Item("connectionString").ToString())
                connection.Open()
                Dim command As New SqlCommand()
                command.Connection = connection

                'UNION
                For i As Integer = 0 To ParaTables.Count - 1

                    If i = ParaTables.Count - 1 Then
                        Sql &= "SELECT * FROM " & ParaTables(i).ToString()

                    Else
                        Sql &= "SELECT * FROM " & ParaTables(i).ToString() & "  UNION ALL "

                    End If

                Next

                command.CommandText = Sql
                Adapter.SelectCommand = command
                Adapter.Fill(ReturnVale, _DataSetName)

            End Using
        Catch ex As SqlException
            MessageBox.Show(ex.Message)
            Throw

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Throw

        End Try

        Return ReturnVale
    End Function



#End Region

#Region "単一のテーブル"

    ''' <summary>
    ''' 単一のテーブルのデータセットを返す
    ''' </summary>
    ''' <returns>UNIONのSQLを実行したデータセット</returns>
    Public Function GetTable(TableName As String) As DataSet
        Dim ReturnVale As New DataSet
        Dim Sql As String = "SELECT * FROM "

        Try
            If TableName.Trim() = String.Empty Then
                Throw New Exception("テーブル名がセットされていません")
            Else
                UniqueDataSetName = TableName

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Throw
        End Try


        Dim Adapter As New SqlDataAdapter
        Try
            Using connection As New SqlConnection(My.Settings.Item("connectionString").ToString())
                connection.Open()
                Dim command As New SqlCommand()
                command.Connection = connection
                command.CommandText = Sql & TableName
                Adapter.SelectCommand = command

                'TODO
                Adapter.Fill(ReturnVale, TableName)
                'Adapter.Fill(ReturnVale, "TableName")
                'Adapter.Fill(ReturnVale)
            End Using
        Catch ex As SqlException
            MessageBox.Show(ex.Message)
            Throw

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Throw

        End Try

        Return ReturnVale
    End Function


#End Region

#End Region

End Class
