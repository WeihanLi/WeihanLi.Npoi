FROM mcr.microsoft.com/dotnet/sdk:8.0
RUN apt-get update \ 
    && apt-get install -y libgdiplus libc6-dev \ 
    && ln -s /usr/lib/libgdiplus.so /usr/lib/gdiplus.dll
