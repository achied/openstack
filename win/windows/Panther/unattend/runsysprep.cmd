cd /d %~dp0

net use * /d /y

certutil -addstore "TrustedPublisher" c:\drivers\setup\rhdriver.cer
::c:\drivers\setup\devcon.exe install c:\drivers\virtio\viostor.inf PCI\VEN_1AF4&DEV_1001&SUBSYS_00021AF4
c:\drivers\setup\devcon.exe dp_add c:\drivers\virtio\viostor.inf
c:\drivers\setup\devcon.exe dp_add c:\drivers\virtio\netkvm.inf

::disabling ipv6 here to make scripts later easier
regedit /S c:\drivers\setup\disableipv6.reg


c:\windows\system32\sysprep\sysprep.exe /oobe /generalize /unattend:%cd%\unattend.xml