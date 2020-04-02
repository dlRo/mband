# Right-click / Send To Copy PS1 file to TXT file
It is often necessary to email a PS1 file, and often files with a ps1 file extension are blocked since it would be so easy to send something malicious that way, so the usual technique is to copy / paste the file and change the suffix of the new file to txt, and that can be emailed.

Here's a method that will allow you to right-click a ps1 file in Windows File Explorer, then Send To... copy-toTxt.bat, and the result is that file as a txt file. A pleasant surprise was the Date modified setting remains as was the original. And in fact ths can be used on any file suffix, vbs being the obvious other use case.

1. Copy Copy-ToTxt.ps1 onto your system
2. Access the Send To folder by typing Win+R and in the Open text field enter shell:sendto then click OK
3. Create Copy-ToTxt.bat file there with these contents, for example:
powershell -NoProfile -ExecutionPolicy Bypass c:\users\username\Documents\WindowsPowerShell\Copy-ToTxt.ps1 -File '%1'
Done! Right-click any ps1 file, Send To \ copy-ToTxt.bat and a txt file will appear.
