#sysprep_hta.ps1
#Last Updated: 7/21/2017
#Decription:
#Run sysprep after deploying an image so that we can move this hard drive to a new computer

Start-Process -FilePath C:\Windows\System32\Sysprep\Sysprep.exe -ArgumentList '/generalize /oobe /shutdown'