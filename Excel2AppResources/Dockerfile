FROM mcr.microsoft.com/dotnet/runtime:7.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /src
COPY ["Excel2AppResources/Excel2AppResources.csproj", "Excel2AppResources/"]
RUN dotnet restore "Excel2AppResources/Excel2AppResources.csproj"
COPY . .
WORKDIR "/src/Excel2AppResources"
RUN dotnet build "Excel2AppResources.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "Excel2AppResources.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Excel2AppResources.dll"]
