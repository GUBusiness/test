Set-Variable uri 'https://kakko116.atlassian.net' -Option Constant
$userName="kakko.1.16.2@gmail.com"
$userPwd="e2008019"
$uri2 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/Book1.xlsx?version=1&modificationDate=1462787267198&api=v2&download=true'
$uri3 = 'https://kakko116.atlassian.net/wiki/download/attachments/655363/vba120.xls?version=1&modificationDate=1462960310042&api=v2&download=true'

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
Start-Sleep -milliseconds 2000
#IE��~
$ie.Quit()

#������
Start-Sleep -milliseconds 2000
$ie=new-object -com InternetExplorer.Application
$ie.visible=$true
$ie.navigate("$uri2")

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

#�_�E�����[�h�̕\����ʂ̃v���Z�X�폜
Get-Process iexplore | Foreach-Object { $_.CloseMainWindow()}


}catch [Exception]{
    $Error
}finally{
    exit
}
