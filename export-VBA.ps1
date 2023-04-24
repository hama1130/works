<#
使用方法：./export-VBA.ps1 .\sample.xlsm
        （エイリアス設定を推奨）
機能　　：引数で指定したファイルからVBAをテキストファイルとして出力する。
備考　　：事前にExcelのオプションから「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」の設定が必要
#>

#引数チェック
if($args.Count -ne 1){
    echo "引数にファイル名を指定してください"
    exit
}

#拡張子チェック
$filename = $args[0]
if(-not ($filename -like "*.xlsm")){
    echo ".xlsmファイルを指定してください"
    exit
}

#絶対パスに変更
$filename = $(resolve-path $filename).ProviderPath

#VBObjectの種類一覧
$hash = @{
    1=".bas";           #1 : 標準モジュール(bas)
    2=".cls";           #2: クラスモジュール(cls)
    3=".frm";           #3: フォーム(.frm)
    11=".cls";          #11: ActivezX (.cls)
    100=".cls"          #100: ドキュメント・シート(.cls)
}

#Excelアプリを開く
$exl =new-object -comobject excel.application
# $exl.visible =$true 
$exl.DisplayAlerts =$False
$exl.EnableEvents = $False

#ブックを開く(読み取り専用）　※３番目がReadOnly引数
$wb=$exl.Workbooks.Open($filename, $null, $true)

#出力先フォルダを作成
$dir_name = "VBA_$($wb.name)".replace(".xlsm","")
mkdir $dir_name | cd

#エクスポート
$vbcmps = $wb.VBProject.VBComponents
 foreach ($tmp in $vbcmps){
    $tmp.Export($(pwd).ProviderPath + "\"+$tmp.name + $hash[$tmp.Type])
    echo "$($tmp.name)をエクスポート"
 }

#Excelを閉じる(保存しない）
$wb.close($False)
$exl.EnableEvents = $True
$exl.Quit()

#終了処理
$exl = $null
[GC]::Collect()

#メッセージ
echo "エクスポートが終了しました"

#元のフォルダに戻る
cd ..