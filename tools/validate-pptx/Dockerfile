# Tier-2 OOXML schema validator (ROADMAP Phase 5).
# Contributors and CI run this image — no .NET SDK on any host.
#
#   docker build -t pptx-validator tools/validate-pptx
#   docker run --rm -v "$PWD:/repo:ro" pptx-validator \
#     --allowlist /repo/tools/validate-pptx/allowlist.json /repo/__tests__/pptx-output

FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS build
WORKDIR /src
COPY validate-pptx.csproj .
RUN dotnet restore
COPY Program.cs .
RUN dotnet publish -c Release -o /app --no-restore

FROM mcr.microsoft.com/dotnet/runtime:8.0-alpine
COPY --from=build /app /app
ENTRYPOINT ["dotnet", "/app/validate-pptx.dll"]
