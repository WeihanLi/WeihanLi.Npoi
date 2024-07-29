[string]$SCRIPT = '.\build.cs'
 
# Install dotnet tool
dotnet tool install --global dotnet-execute

Write-Host "dotnet-exec $SCRIPT --args $ARGS" -ForegroundColor GREEN
 
dotnet-exec $SCRIPT --args $ARGS