<#
使用方法：./get-SQL.ps1 "DB.dbo.testprocedure"
概要　　：DWHサーバーに接続し、オブジェクトの定義分をコンソールに出力する
特徴　　：オブジェクト名は引数から与える。
　　　　　リダイレクトによりテキストファイルへの出力が可能。
　　　　　SQL Server Management Studioから作成したスクリプトと同様の形式に調整済み。　
引数　　　：オブジェクト名
#>

# 引数定義（オブジェクト名）
Param([Parameter(Mandatory = $true)][String]$object_name)

#引数を分解
$ary = $object_name.split(".")
$dbname = $ary[0]
$schema = $ary[1]
$name = $ary[2]

#ダブルクオーテーションがあるとエラーになるので消去
$object_name = $object_name.Replace('"','')

#定数
$datasource = "localhost" #SQLサーバーのホスト
$database_name = $dbname #データベース名
$SQLstring = "select definition from $dbname.sys.sql_modules where object_id=object_id('$object_name')" #SQL文字列1
$SQLstring2 = "select type from $dbname.sys.objects where object_id=object_id('$object_name')" #SQL文字列2

# SqlConnectionStringBuilder を使用してSQL接続の設定を保存する
$ConnectionString = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder

#接続
$ConnectionString['Data Source'] = $datasource #SQLServerのホストを指定
$ConnectionString['Initial Catalog'] = $database_name #データベースを指定
$ConnectionString['Integrated Security'] = "TRUE" #Windows統合認証を利用する場合は"TRUE"

# SQL文の文字列を設定する
$SQLQuery = $SQLstring

# DataTableを利用してSQL実行結果を一時格納
$resultsDataTable = New-Object System.Data.DataTable

# SQLConnection、SQLCommandを設定する
$SqlConnection = New-Object System.Data.SQLClient.SQLConnection($ConnectionString)
$SqlCommand = New-Object System.Data.SQLClient.SQLCommand($SQLQuery, $SqlConnection)

# データベースへ接続
$SqlConnection.Open()

# ExecuteReaderを実行してDataTableにデータを格納
$resultsDataTable.Load($SqlCommand.ExecuteReader())

# データベース接続解除
$SqlConnection.Close()

#変数に格納
$definition = $resultsDataTable.definition

#2 種類を取得
# SQL文の文字列を設定する
$SQLQuery = $SQLstring2

# 格納先を定義
$resultsDataTable = New-Object System.Data.DataTable

# SQLConnection、SQLCommandを設定する
$SqlConnection = New-Object System.Data.SQLClient.SQLConnection($ConnectionString)
$SqlCommand = New-Object System.Data.SQLClient.SQLCommand($SQLQuery, $SqlConnection)

# データベースへ接続
$SqlConnection.Open()
$resultsDataTable.Load($SqlCommand.ExecuteReader())
$SqlConnection.Close()

# 結果を変数に格納
$type = $resultsDataTable.type

#オブジェクトの種類の変換用ハッシュテーブル
$hash = @{
"P "="StoredProcedure";
"TF"="UserDefinedFunction";
"U "="Table";
"V "="View";
"FN"="UserDefinedFunction";
"IF"="UserDefinedFunction";
}

#代入用の変数を定義
$datetime = date -Format "g"
$type_def = $hash[$type]

#出力テキスト作成
$preamble =@"
USE [$dbname]
GO

/****** Object:  $type_def [$schema].[$name]    Script Date: $datetime ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


"@

$postscript = @"

GO


"@

#文字列を連結
$output = $preamble + $definition + $postscript

#出力
echo $output