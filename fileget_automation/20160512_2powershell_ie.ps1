Set-Variable uri 'https://kakko116.atlassian.net' -Option Constant
$userName="kakko.1.16.2@gmail.com"
$userPwd="e2008019"
$uri2 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/Book1.xlsx?version=1&modificationDate=1462787267198&api=v2&download=true'
$uri3 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/vba120.xls?version=1&modificationDate=1462960310042&api=v2&download=true'
$logout = 'https://kakko116.atlassian.net/logout'
$workUrl='https://kakko116.atlassian.net/wiki/display/WOR'
#DL�t�@�C������΃p�X
$File_Base1 = 'C:\Users\ari\Desktop\fileget_automation\DL�t�@�C���y���_A�z.txt'
$Base1_URL = 'C:\Users\ari\Desktop\fileget_automation\Base1List.txt'


#�g���b�v
try{
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri")


#�҂�����
While($ie.Busy){
	start-sleep -milliseconds 100
}
#DOM��ID���Ȃǂ͓K�X�ύX���Ă��������B
$doc=$ie.document
$dom_userNAME=$doc.getElementById("username")
$dom_userNAME.value=$userName
$dom_userPWD=$doc.getElementById("password")
$dom_userPWD.value=$userPwd

$btn=$doc.getElementByID("login")
$btn.click()

#�҂�����
While($ie.Busy){
	start-sleep -milliseconds 100
}

#Book1.txt�̃t�@�C�������N�擾
$ie.navigate("$workUrl")

#�҂�����
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
#IE��~
$ie.Quit()

<#
#������
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
#$ie.navigate("$uri2")
$ie.navigate("$f1")

#���O�C������5�b�҂�
Start-Sleep -milliseconds 5000
#���O�C����
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#�|�b�v�A�b�v�\���҂�10�b
Start-Sleep -milliseconds 10000
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
#IE��~
#$ie.Quit()

#���J��Ԃ�
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri3")

#���O�C������5�b�҂�
Start-Sleep -milliseconds 5000
#���O�C����
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#�|�b�v�A�b�v�\���҂�10�b
Start-Sleep -milliseconds 10000
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
#IE��~
#$ie.Quit()
#���J��Ԃ�


#logout����
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$logout")
#�҂�����
While($ie.Busy){
	start-sleep -milliseconds 1000
}
$doc=$ie.document
$btn=$doc.getElementByID("logout")
$btn.click()


#�_�E�����[�h�̕\����ʂ̃v���Z�X�폜
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}

#>
}catch [Exception]{
    $Error
}finally{
    exit
}
