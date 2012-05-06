#---------------------------------
#�摜���W�X�N���v�g
# image search & download script (Google Search)
# kazunori.kimura.js@gmail.com
#---------------------------------
# �A�Z���u���ǂݍ���
[void]([Reflection.Assembly]::LoadWithPartialName("System.Web"))
[void]([Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions"))

#�f�o�b�O�o�͂�L���ɂ���
$DebugPreference = "Continue" #"SilentlyContinue"
#�x���o�͂�L���ɂ���
$WarningPreference = "Continue"

#�擾���� (4 or 8)
$RESULT_SIZE = 8
#�ő匟����
$MAX_SEARCH_COUNT = 3

#Google Search URL
$URL="https://ajax.googleapis.com/ajax/services/search/images?"
#�����I�v�V����
$option = @{
	"v"="1.0";
	"searchType"="image";
	"rsz"=8;
	"q"="";
}


#
# ����URL��g�ݗ��Ă�
#
function Get-SearchURL($word, $start=0, $rsz=4){
	$ret = ""
	try{
		#�L�[���[�h���G���R�[�h���ăZ�b�g
		$option["q"] = [Web.HttpUtility]::UrlEncode($word)
		
		#�J�n�ʒu
		$option["start"] = $start
		
		#URL�g��
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
# Google����JSON�`���Ō������ʂ�GET����B
#
function Get-SearchResult($url){
	$ret = ""
	try{
		$req = [Net.HttpWebRequest]::Create($url)
		$req.Method = "GET"
		#UserAgent�w��K�v�H
		
		#�������ʎ擾
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
# ���g�p���Ď擾����JSON���p�[�X����B
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
# �摜URL���擾���A�_�E�����[�h���X�g���o�͂���B
#
function Get-DownloadList($keyword){
#�������ʂ̌`
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
			#�������s�񐔂��K��񐔈ȏ�ɂȂ����ꍇ�͏I��
			break
		}
		#��������
		$json = Get-SearchResult (Get-SearchURL $keyword $i $RESULT_SIZE)
		#���ʂ��p�[�X
		$res = Parse-JSON $json
		#�������ʌ���
		$resultCount = $res.responseData.cursor.estimatedResultCount
		#���ʂ��o��
		foreach($item in $res.responseData.results){
			$item["url"]
		}
		
		$loopCount++ #�������s��
		Write-Debug ("������= $loopCount`t��������= $resultCount")
	}
}

#
# �摜���_�E�����[�h����
#
function Download-File($list, $folder=".\"){
	try{
		$client = New-Object System.Net.WebClient
		
		foreach( $url in Get-Content $list ){
			$uri = New-Object System.Uri($url)
			#�t�@�C����
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

Get-DownloadList "�q���g仐�" > .\download.txt
Download-File .\download.txt "C:\Users\kaz\Pictures\test"