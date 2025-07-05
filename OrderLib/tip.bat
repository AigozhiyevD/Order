dotnet publish -c Release -r win-x64 --self-contained -p:PublishReadyToRun=true 
del bin\Release\net8.0\win-x64\*.*