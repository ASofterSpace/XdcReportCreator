REM set /p PhpPath=<php-path.config

REM %PhpPath%\php -f test.php

REM pause

REM Actually, let PHP call Java when it is necessary...

java -classpath "bin" -Xms16m -Xmx1024m com.asofterspace.xdcReportCreator.Main %*

pause
