@ECHO OFF

setlocal EnableDelayedExpansion
title Windows KMS激活脚本

%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

for /f "delims=" %%i in ('REG query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName') do (set osver=%%i)

SET osver=%osver:~29%
SET kmsserver=hak5.ga

:: 备用服务器
:: SET kmsserver=kms.lotro.cc

ECHO [+] 当前系统为：%osver%&ECHO.
ECHO [+] 激活服务器：%kmsserver%&ECHO.

:: Windows 桌面系统版激活码
IF "%osver%" == "Windows 10 Enterprise" (SET key=NPPR9-FWDCX-D2C8J-H872K-2YT43 & GOTO :Active)
IF "%osver%" == "Windows 10 Enterprise N" (SET key=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 & GOTO :Active)
IF "%osver%" == "Windows 10 Professional" (SET key=W269N-WFGWX-YVC9B-4J6C9-T83GX & GOTO :Active)
IF "%osver%" == "Windows 10 Professional N" (SET key=MH37W-N47XK-V7XM9-C7227-GCQG9 & GOTO :Active)
IF "%osver%" == "Windows 10 Enterprise 2016 LTSB" (SET key=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ & GOTO :Active)
IF "%osver%" == "Windows 10 Enterprise 2015 LTSB" (SET key=WNMTR-4C88C-JK8YV-HQ7T2-76DF9 & GOTO :Active)

IF "%osver%" == "Windows 8.1 Enterprise" (SET key=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 & GOTO :Active)
IF "%osver%" == "Windows 8.1 Enterprise N" (SET key=TT4HM-HN7YT-62K67-RGRQJ-JFFXW & GOTO :Active)
IF "%osver%" == "Windows 8.1 Professional" (SET key=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 & GOTO :Active)
IF "%osver%" == "Windows 8.1 Professional N" (SET key=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY & GOTO :Active)

IF "%osver%" == "Windows 8 Enterprise" (SET key=32JNW-9KQ84-P47T8-D8GGY-CWCK7 & GOTO :Active)
IF "%osver%" == "Windows 8 Enterprise N" (SET key=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT & GOTO :Active)
IF "%osver%" == "Windows 8 Professional" (SET key=NG4HW-VH26C-733KW-K6F98-J8CK4 & GOTO :Active)
IF "%osver%" == "Windows 8 Professional N" (SET key=XCVCF-2NXM9-723PB-MHCB7-2RYQQ & GOTO :Active)

IF "%osver%" == "Windows 7 Enterprise" (SET key=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH & GOTO :Active)
IF "%osver%" == "Windows 7 Enterprise N" (SET key=YDRBP-3D83W-TY26F-D46B2-XCKRJ & GOTO :Active)
IF "%osver%" == "Windows 7 Enterprise E" (SET key=C29WB-22CC8-VJ326-GHFJW-H9DH4 & GOTO :Active)
IF "%osver%" == "Windows 7 Professional" (SET key=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 & GOTO :Active)
IF "%osver%" == "Windows 7 Professional N" (SET key=NMRPKT-YTG23-K7D7T-X2JMM-QY7MG & GOTO :Active)
IF "%osver%" == "Windows 7 Professional E" (SET key=W82YF-2Q76Y-63HXB-FGJG9-GF7QX & GOTO :Active)

IF "%osver%" == "Windows Vista Business" (SET key=YFKBB-PQJJV-G996G-VWGXY-2V3X8 & GOTO :Active)
IF "%osver%" == "Windows Vista Business N" (SET key=HMBQG-8H2RH-C77VX-27R82-VMQBT & GOTO :Active)
IF "%osver%" == "Windows Vista Enterprise" (SET key=VKK3X-68KWM-X2YGT-QR4M6-4BWMV & GOTO :Active)
IF "%osver%" == "Windows Vista Enterprise N" (SET key=VTC42-BM838-43QHV-84HX6-XJXKV & GOTO :Active)

:: Windows 服务器版激活码
IF "%osver%" == "Windows Server 2016 Datacenter" (SET key=CB7KF-BWN84-R7R2Y-793K2-8XDDG & GOTO :Active)
IF "%osver%" == "Windows Server 2016 Standard" (SET key=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY & GOTO :Active)
IF "%osver%" == "Windows Server 2016 Essentials" (SET key=JCKRF-N37P4-C2D82-9YXRT-4M63B & GOTO :Active)

IF "%osver%" == "Windows Server 2012 R2 Datacenter" (SET key=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9 & GOTO :Active)
IF "%osver%" == "Windows Server 2012 R2 Essentials" (SET key=KNC87-3J2TX-XB4WP-VCPJV-M4FWM & GOTO :Active)
IF "%osver%" == "Windows Server 2012 R2 Server Standard" (SET key=D2N9P-3P6X9-2R39C-7RTCD-MDVJX & GOTO :Active)
IF "%osver%" == "Windows Server 2012" (SET key=BN3D2-R7TKB-3YPBD-8DRP2-27GG4 & GOTO :Active)
IF "%osver%" == "Windows Server 2012 N" (SET key=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY & GOTO :Active)
IF "%osver%" == "Windows Server 2012 Single Language" (SET key=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ & GOTO :Active)
IF "%osver%" == "Windows Server 2012 Country Specific" (SET key=4K36P-JN4VD-GDC6V-KDT89-DYFKP & GOTO :Active)
IF "%osver%" == "Windows Server 2012 Server Standard" (SET key=XC9B7-NBPP2-83J2H-RHMBY-92BT4 & GOTO :Active)
IF "%osver%" == "Windows Server 2012 MultiPoint Standard" (SET key=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ & GOTO :Active)
IF "%osver%" == "Windows Server 2012 MultiPoint Premium" (SET key=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G & GOTO :Active)
IF "%osver%" == "Windows Server 2012 Datacenter" (SET key=48HP8-DN98B-MYWDG-T2DCC-8W83P & GOTO :Active)

IF "%osver%" == "Windows Server 2008 R2 Standard" (SET key=YC6KT-GKW9T-YTKYR-T4X34-R7VHC & GOTO :Active)
IF "%osver%" == "Windows Server 2008 R2 Enterprise" (SET key=489J6-VHDMP-X63PK-3K798-CPX3Y & GOTO :Active)
IF "%osver%" == "Windows Server 2008 R2 Datacenter" (SET key=74YFP-3QFB3-KQT8W-PMXWJ-7M648 & GOTO :Active)
IF "%osver%" == "Windows Server 2008 R2 Web" (SET key=6TPJF-RBVHG-WBW2R-86QPH-6RTM4 & GOTO :Active)
IF "%osver%" == "Windows Server 2008 R2 HPC edition" (SET key=TT8MH-CG224-D3D7Q-498W2-9QCTX & GOTO :Active)
IF "%osver%" == "Windows Server 2008 R2 for Itanium-based Systems" (SET key=GT63C-RJFQ3-4GMB6-BRFB9-CB83V & GOTO :Active)

IF "%osver%" == "Windows Web Server 2008" (SET key=WYR28-R7TFJ-3X2YQ-YCY4H-M249D & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Standard" (SET key=TM24T-X9RMF-VWXK6-X8JC9-BFGM2 & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Standard without Hyper-V" (SET key=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Enterprise" (SET key=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Enterprise without Hyper-V" (SET key=39BXF-X8Q23-P2WWT-38T2F-G3FPG & GOTO :Active)
IF "%osver%" == "Windows Server 2008 HPC" (SET key=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Datacenter" (SET key=7M67G-PC374-GR742-YH8V4-TCBY3 & GOTO :Active)
IF "%osver%" == "Windows Server 2008 Datacenter without Hyper-V" (SET key=22XQ2-VRXRG-P8D42-K34TD-G3QQC & GOTO :Active)
IF "%osver%" == "Windows Server 2008 for Itanium-Based Systems" (SET key=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK & GOTO :Active) ELSE GOTO :Error


:: 激活程序
:Active
ECHO [+] 所使用的激活码为：%key%&ECHO.
ECHO [+] 正在激活，请注意点击每个弹窗的确定按钮&ECHO.

slmgr /upk
slmgr /ipk %key%
slmgr /skms %kmsserver%
slmgr /ato
slmgr /dlv

IF ERRORLEVEL 0 (ECHO [+] 恭喜你，激活成功
) ELSE (ECHO [-] 激活失败，请重试)

ECHO.&PAUSE&GOTO:EOF

:: 不支持激活的系统提示
:Error
ECHO.[-] 该系统暂不支持KMS激活
ECHO.&PAUSE&GOTO:EOF
