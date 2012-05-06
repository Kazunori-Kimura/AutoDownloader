#---------------------------------
#画像収集スクリプト
# image search & download script (Google Search)
# kazunori.kimura.js@gmail.com
#---------------------------------
# アセンブリ読み込み
[void]([Reflection.Assembly]::LoadWithPartialName("System.Web"))
[void]([Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions"))

#デバッグ出力を有効にする
$DebugPreference = "Continue" #"SilentlyContinue"
#警告出力を有効にする
$WarningPreference = "Continue"

#取得件数 (4 or 8)
$RESULT_SIZE = 8
#最大検索回数
$MAX_SEARCH_COUNT = 3

#Google Search URL
$URL="https://ajax.googleapis.com/ajax/services/search/images?"
#検索オプション
$option = @{
	"v"="1.0";
	"searchType"="image";
	"rsz"=8;
	"q"="";
}


#
# 検索URLを組み立てる
#
function Get-SearchURL($word, $start=0, $rsz=4){
	$ret = ""
	try{
		#キーワードをエンコードしてセット
		$option["q"] = [Web.HttpUtility]::UrlEncode($word)
		
		#開始位置
		$option["start"] = $start
		
		#URL組立
		$prm = ""
		foreach($k in $option.keys){
			if( $prm.length -gt 0 ){
				$prm += "&"
			}
			$prm += $k + "=" + $option[$k]
		}
		
		$ret = $URL + $prm
	}catch [Exception]{
		throw $Error[0]
	}
	return $ret
}

#
# GoogleからJSON形式で検索結果をGETする。
#
function Get-SearchResult($url){
	$ret = ""
	try{
		$req = [Net.HttpWebRequest]::Create($url)
		$req.Method = "GET"
		#UserAgent指定必要？
		
		#検索結果取得
		$res = $req.GetResponse()
		$reader = New-Object IO.StreamReader($res.GetResponseStream(), $res.ContentEncoding)
		$ret = $reader.ReadToEnd()
		
		$reader.Close()
		$res.Close()
		
	}catch [Exception]{
		throw $Error[0]
	}
	return $ret
}

#
# System.Web.Script.Serialization.JavaScriptSerializer
# を使用して取得したJSONをパースする。
# [http://msdn.microsoft.com/ja-jp/library/system.web.script.serialization.javascriptserializer.aspx]
#
function Parse-JSON($json){
	$ret = $null
	try{
		$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		$ret = $serializer.DeserializeObject($json)
	}catch [Exception]{
		throw $Error[0]
	}
	return $ret
}

#
# 画像URLを取得し、ダウンロードリストを出力する。
#
function Get-DownloadList($keyword){
#検索結果の形
#	responseData.results[]
#		.titleNoFormatting
#		.url
#		.imageId
#		.height
#		.width
	#$fields = @("titleNoFormatting", "url", "imageId")

	$resultCount = $RESULT_SIZE
	$loopCount = 0
	for( $i=0; $i -lt $resultCount; $i += $RESULT_SIZE ){
		if( $loopCount -ge $MAX_SEARCH_COUNT){
			#検索実行回数が規定回数以上になった場合は終了
			break
		}
		#検索処理
		$json = Get-SearchResult (Get-SearchURL $keyword $i $RESULT_SIZE)
		#結果をパース
		$res = Parse-JSON $json
		#検索結果件数
		$resultCount = $res.responseData.cursor.estimatedResultCount
		#結果を出力
		foreach($item in $res.responseData.results){
			$item["url"]
		}
		
		$loopCount++ #検索実行回数
		Write-Debug ("処理回数= $loopCount`t検索結果= $resultCount")
	}
}

#
# 画像をダウンロードする
#
function Download-File($list, $folder=".\"){
	try{
		$client = New-Object System.Net.WebClient
		
		foreach( $url in Get-Content $list ){
			$uri = New-Object System.Uri($url)
			#ファイル名
			$file = Split-Path $uri.AbsolutePath -Leaf
			try{
				$client.DownloadFile($uri, (Join-Path $folder $file))
				Write-Debug ("Downloaded File= $file")
			}catch [Exception]{
				Write-Warning ("URL= $url`t" + $Error[0].ToString())
			}
		}
	}catch [Exception]{
		throw $Error[0]
	}
}

Get-DownloadList "牧瀬紅莉栖" > .\download.txt
Download-File .\download.txt "C:\Users\kaz\Pictures\test"