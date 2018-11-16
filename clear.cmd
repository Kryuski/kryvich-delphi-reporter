del /f /s *.~* *.bak *.dcu *.dsk *.identcache *.local *.map *.drc *.exe *.stat *.otares
rd /s /q __history
rd /s /q __recovery
rd /s /q Demo\__history
rd /s /q Demo\__recovery
del /q Demo\Output\*.*
