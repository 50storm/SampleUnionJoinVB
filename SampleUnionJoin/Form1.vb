Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Linq

Imports SampleUnionJoin.QcOut


Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'クラスファイルを使う側



        Dim table1 As String = Me.TextBox1.Text
        Dim table2 As String = Me.TextBox2.Text
        Dim DataSetName As String = "TableNameCommon"


        Dim QCOUT As New QcOut

        Console.WriteLine(Now)

        Dim Dst As DataSet

        Try

            Dst = QCOUT.GetCommonTables(DataSetName.Trim(), table1, table2)



        Catch ex As Exception
            MessageBox.Show("例外メッセージ：" & ex.Message)
            MessageBox.Show("例外トレース：" & ex.StackTrace)


            Exit Sub

        End Try


        'レコードがあるかチェック
        Dim UnionTable As DataTable = Dst.Tables(QCOUT.CommonDataSetName)
        If UnionTable.Columns.Count = 0 Then
            MessageBox.Show("テーブルに１レコードもありませんでした。")
        End If

        'DataTableに取り出す場合
        For Each table As DataTable In Dst.Tables
            Dim foundRows() As DataRow

            'SQLのWHERE区に相当する文を書く
            Dim expression As String = "indexnum>=67"
            foundRows = table.Select(expression)

            'データをすべて取り出す場合
            Console.WriteLine("/////////////////////////////////////////////")

            For i As Integer = 0 To foundRows.GetUpperBound(0)
                Console.WriteLine(foundRows(i)(0))
                For Each record As Object In foundRows(i).ItemArray
                    '一つのレコードのカラムを取り出す場合
                    Console.WriteLine(record)
                Next
            Next i

            Console.WriteLine("/////////////////////////////////////////////")


            For Each row As DataRow In table.Rows


                Dim index As Integer = Integer.Parse(row.Item("indexnum").ToString)
                Dim escort As String = row.Item("escortnum").ToString

                If index >= 67 Then


                    Console.WriteLine(index)
                    Console.WriteLine(escort)
                End If

            Next

            Console.WriteLine("/////////////////////////////////////////////")

        Next



        'LINQを使って取り出す場合
        Dim query() = (From x In UnionTable.AsEnumerable() Where x.Field(Of String)("product") = "ASR55-FSC=" Select x.Field(Of String)("serialnum")).ToArray()
        For Each record As String In query
            Console.WriteLine(record)
        Next


        'LINQサンプル２
        Dim query2 = (From x In UnionTable.AsEnumerable() Where x.Field(Of Integer)("indexnum") >= 67)

        Dim dt_asndate As DateTime
        Dim dt_endinspection As DateTime
        Dim ts As TimeSpan


        For Each record In query2

            Console.WriteLine(record.Field(Of Nullable(Of Date))("asndate").ToString())


            Console.WriteLine(record.Field(Of Nullable(Of Date))("endinspection").ToString())
            dt_asndate = record.Field(Of Nullable(Of Date))("asndate")
            dt_endinspection = record.Field(Of Nullable(Of Date))("endinspection")
            ts = dt_endinspection.Subtract(dt_asndate)

            Console.WriteLine(ts)


        Next




        Console.WriteLine(Now)



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim QCOUT As New QcOut
        Dim Dst2 As DataSet
        Dim table As String = Me.TextBox3.Text.Trim()
        'Dim DataSetName As String = table


        Dst2 = QCOUT.GetTable(table)

        If Dst2.Tables.Count = 0 Then
            'テーブルが取得できませんでした。
            MessageBox.Show("テーブルが取得できませんでした。")
        Else

            Dim Table2 As DataTable = Dst2.Tables(QCOUT.UniqueDataSetName)
            'Dim Table2 As DataTable = Dst2.Tables(DataSetName)
            Dim foundRows() As DataRow = Table2.Select("indexnum >= 67 and indexnum <= 100")
            'データをすべて取り出す場合
            Console.WriteLine("/////////////////////////////////////////////")

            For i As Integer = 0 To foundRows.GetUpperBound(0)
                Console.WriteLine(foundRows(i)(0))
                For Each record As Object In foundRows(i).ItemArray
                    '一つのレコードのカラムを取り出す場合
                    Console.WriteLine(record)
                Next
            Next i

            Console.WriteLine("/////////////////////////////////////////////")



            'LINQサンプル２
            Dim query2 = (From x In Table2.AsEnumerable()
                          Where x.Field(Of Integer)("indexnum") >= 67 And x.Field(Of Integer)("indexnum") <= 100)




            Dim dt_asndate As DateTime
            Dim dt_endinspection As DateTime
            Dim ts As TimeSpan


            For Each record In query2

                Console.WriteLine(record.Field(Of Nullable(Of Date))("asndate").ToString())


                Console.WriteLine(record.Field(Of Nullable(Of Date))("endinspection").ToString())
                dt_asndate = record.Field(Of Nullable(Of Date))("asndate")
                dt_endinspection = record.Field(Of Nullable(Of Date))("endinspection")
                ts = dt_endinspection.Subtract(dt_asndate)

                Console.WriteLine(ts)


            Next

        End If

    End Sub
End Class
