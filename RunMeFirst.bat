:: Registers the "IShellFolder Extended Type Library v1.2"
:: (ISHF_Ex.tlb), the "pToolTip" (pTooltip.dll) and the
:: "RepControls" (RepControls.ocx) which are all used by 
:: RepShell

:: Assumes all files are in current folder, or you will have
:: register them manually. This won't be neccesary when 
:: setup is implemented

@Echo Off
start regsvr32 ISHF_Ex.tlb
start regsvr32 pToolTip.dll
start regsvr32 RepControls.ocx