# Imagen base de runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 8080
ENV ASPNETCORE_URLS=http://+:8080

# Imagen de build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copiamos los .csproj de cada proyecto
COPY ["ModuloWeb1/ModuloWeb1.csproj", "ModuloWeb1/"]
COPY ["ModuloWeb.BROKER/ModuloWeb.BROKER.csproj", "ModuloWeb.BROKER/"]
COPY ["ModuloWeb.MANAGER/ModuloWeb.MANAGER.csproj", "ModuloWeb.MANAGER/"]
COPY ["ModuloWeb.ENTITIES/ModuloWeb.ENTITIES.csproj", "ModuloWeb.ENTITIES/"]

# Restauramos dependencias del proyecto web
RUN dotnet restore "ModuloWeb1/ModuloWeb1.csproj"

# Copiamos el resto del código
COPY . .

# Publicamos en modo Release
WORKDIR "/src/ModuloWeb1"
RUN dotnet publish "ModuloWeb1.csproj" -c Release -o /app/publish

# Imagen final
FROM base AS final
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "ModuloWeb1.dll"]
