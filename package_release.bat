set config="Debug"
set VERSION=1.0.1-Beta
set ZIPFILE=dssplugin-DSSExcel.zip
Xcopy DSSExcelImport\bin\%config% distribution\dotnet\DSSExcelImport /e /h /c /i /s /y
Xcopy DSSExcelExport\bin\%config% distribution\dotnet\DSSExcelExport /e /h /c /i /s /y

cd distribution
del %ZIPFILE%
7z a -r dssplugin-DSSExcel.zip dotnet\*

::  dssexcel 'mil.army.usace.hec:dssplugin-DSSExcel:1.0-Beta@zip'

C:\Programs\apache-maven-3.8.5\bin\mvn deploy:deploy-file -DgroupId=mil.army.usace.hec -DartifactId=dssplugin-DSSExcel -Dversion=%VERSION% -DgeneratePom=true -Dfile=%ZIPFILE% -DrepositoryId=nexus -Dpackaging=zip -Durl=https://www.hec.usace.army.mil/nexus/repository/maven-releases/


cd ..
