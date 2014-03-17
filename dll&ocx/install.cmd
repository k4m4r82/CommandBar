cls
copy SSubTmr6.dll %systemroot%\system32
copy vbalCmdBar6.ocx %systemroot%\system32

regsvr32 /s %systemroot%\system32\SSubTmr6.dll
regsvr32 /s %systemroot%\system32\vbalCmdBar6.ocx