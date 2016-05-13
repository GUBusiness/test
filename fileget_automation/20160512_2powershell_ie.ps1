Set-Variable uri 'https://kakko116.atlassian.net' -Option Constant
$userName="kakko.1.16.2@gmail.com"
$userPwd="e2008019"
$uri2 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/Book1.xlsx?version=1&modificationDate=1462787267198&api=v2&download=true'
$uri3 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/vba120.xls?version=1&modificationDate=1462960310042&api=v2&download=true'
$logout = 'https://kakko116.atlassian.net/logout'
$workUrl='https://kakko116.atlassian.net/wiki/display/WOR'
#DLファイル名絶対パス
$File_Base1 = 'C:\Users\ari\Desktop\fileget_automation\DLファイル【拠点A】.txt'
$Base1_URL = 'C:\Users\ari\Desktop\fileget_automation\Base1List.txt'


#トラップ
try{
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri")


#待ち時間
While($ie.Busy){
	start-sleep -milliseconds 100
}
#DOMのID名などは適宜変更してください。
$doc=$ie.document
$dom_userNAME=$doc.getElementById("username")
$dom_userNAME.value=$userName
$dom_userPWD=$doc.getElementById("password")
$dom_userPWD.value=$userPwd

$btn=$doc.getElementByID("login")
$btn.click()

#待ち時間
While($ie.Busy){
	start-sleep -milliseconds 100
}

#Book1.txtのファイルリンク取得
$ie.navigate("$workUrl")

#待ち時間
While($ie.Busy){
	start-sleep -milliseconds 100
}


$f = 0;
$base1 = get-content $File_Base1
Foreach($x in $base1){
write-host $x
$element = $ie.Document.getElementsByTagName('A') |
                where-object {
                    $_.innerText -eq $x
                }
$f = $element.href
$f = $f + '&download=true'
$f >> $Base1_URL


}
Start-Sleep -milliseconds 2000
#IE停止
$ie.Quit()

<#
#お試し
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
#$ie.navigate("$uri2")
$ie.navigate("$f1")

#ログイン処理5秒待ち
Start-Sleep -milliseconds 5000
#ログイン後
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#ポップアップ表示待ち10秒
Start-Sleep -milliseconds 10000
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
#IE停止
#$ie.Quit()

#↓繰り返し
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri3")

#ログイン処理5秒待ち
Start-Sleep -milliseconds 5000
#ログイン後
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#ポップアップ表示待ち10秒
Start-Sleep -milliseconds 10000
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
#IE停止
#$ie.Quit()
#↑繰り返し


#logout処理
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$logout")
#待ち時間
While($ie.Busy){
	start-sleep -milliseconds 1000
}
$doc=$ie.document
$btn=$doc.getElementByID("logout")
$btn.click()


#ダウンロードの表示画面のプロセス削除
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}

#>
}catch [Exception]{
    $Error
}finally{
    exit
}
