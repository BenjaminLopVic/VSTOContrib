msbuild src\VSTOContrib.Core\VSTOContrib.Core.csproj /p:Configuration=Release
msbuild src\VSTOContrib.Word\VSTOContrib.Word.csproj /p:Configuration=Release
msbuild src\VSTOContrib.Autofac\VSTOContrib.Autofac.csproj /p:Configuration=Release

nuget pack src\VSTOContrib.Core\VSTOContrib.Core.csproj -Symbols -IncludeReferencedProjects -SymbolPackageFormat snupkg -OutputDirectory .\Nugets -Properties Configuration=Release
nuget pack src\VSTOContrib.Word\VSTOContrib.Word.csproj -Symbols -IncludeReferencedProjects -SymbolPackageFormat snupkg -OutputDirectory .\Nugets -Properties Configuration=Release
nuget pack src\VSTOContrib.Autofac\VSTOContrib.Autofac.csproj -Symbols -IncludeReferencedProjects -SymbolPackageFormat snupkg -OutputDirectory .\Nugets -Properties Configuration=Release