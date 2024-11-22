# ExportToWordBlazorWASM
Función para sobreescribir una plantilla de Microsoft Word desde Blazor WebAssembly.

#Notas
- Fue utilizado en .NET 8.0

#Librerías utilizadas:
- DocumentFormat.OpenXml
- System.IO
- Blazored.LocalStorage

#Modo de uso:
1. Instalar las librerías:
    - dotnet add package DocumentFormat.OpenXml
    - dotnet add package Blazored.LocalStorage
2. Incluirlas en el archivo _Imports.razor o utilizarlas directamente en el archivo que se van a utilizar
    - @using DocumentFormat.OpenXml.Packaging;
    - @using DocumentFormat.OpenXml.Wordprocessing;
    - @using System.IO;
    - @using Blazored.LocalStorage
3. Añadir el archivo download.js a la carpeta "js" e incluirlo en el index.html que está en la carpeta wwwroot del proyecto
    - <script src="js/download.js"></script>
4. Incluir el servicio de HttpClient, servicio "IWordService" y su implementación "WordService" y el servicio de "LocalStorage"  en  en el Program.cs
    - builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
    - builder.Services.AddScoped<IWordService, WordService>();
    - builder.Services.AddBlazoredLocalStorage();
6. Inyectar los servicios IWordService, IJSRuntime, HttpClient, IlocalStorage en el archivo razor donde se desea interactuar con el documento Word
    - @inject IJSRuntime JS
    - @inject ILocalStorageService localStorage
    - @inject HttpClient Http
    - @inject IWordService WordService

#Recomendaciones
- El código razor que que se encuentra en el repositorio es un ejemplo para demostrar qué parametros puede recibir la función GenerateWordDocument en el IWordService para que lo usen de guía
- El archivo de Word recomiendo que tenga la extensión .docx, no lo probé con otras extensiones
- Creen una carpera dentro de wwwroot para almacenar sus plantillas, en mi caso la llamé templates
- Verifiquen los namespaces ya que los namespaces que están en estos archivos son los de mi proyecto no es recomendable que usen los mismos que los míos
    
