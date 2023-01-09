set config="Debug"
set VERSION=1.0.2-Beta
set ZIPFILE=dssplugin-DSSExcel.zip
Xcopy DSSExcel\bin\%config% distribution\dotnet\DSSExcel /e /h /c /i /s /y

cd distribution
del %ZIPFILE%
7z a -r  %ZIPFILE% dotnet\*

mvn deploy:deploy-file -DgroupId=mil.army.usace.hec -DartifactId=dssplugin-DSSExcel -Dversion=%VERSION% -DgeneratePom=true -Dfile=%ZIPFILE% -DrepositoryId=nexus -Dpackaging=zip -Durl=https://www.hec.usace.army.mil/nexus/repository/maven-releases/


cd ..
