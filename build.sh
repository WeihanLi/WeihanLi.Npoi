#!/bin/sh
SCRIPT='./build.cs'

# Install tool
dotnet tool install --global dotnet-execute
export PATH="$PATH:$HOME/.dotnet/tools"

echo "dotnet-exec $SCRIPT --args=$@"

dotnet-exec $SCRIPT --args="$@"