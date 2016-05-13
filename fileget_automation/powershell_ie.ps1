Set-Variable uri 'https://kakko116.atlassian.net' -Option Constant
$userName="kakko.1.16.2@gmail.com"
$userPwd="e2008019"
$uri2 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/Book1.xlsx?version=1&modificationDate=1462787267198&api=v2&download=true'
$uri3 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/vba120.xls?version=1&modificationDate=1462960310042&api=v2&download=true'

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
Start-Sleep -milliseconds 2000
#IE停止
$ie.Quit()

#お試し
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri2")

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

#ダウンロードの表示画面のプロセス削除
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}


}catch [Exception]{
    $Error
}finally{
    exit
}
