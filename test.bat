set /p PhpPath=<php-path.config

%PhpPath%\php -f test.php

pause

REM Actually, let PHP call Java when it is necessary...

java -classpath "bin" -Xms16m -Xmx1024m com.asofterspace.xdcReportCreator.Main %*

pause
