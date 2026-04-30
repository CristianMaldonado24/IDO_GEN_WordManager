# IDO_GEN_WordManager

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

Herramienta de escritorio WPF para la gestión de estructura de documentos Word. Permite cargar un archivo `.docx`, visualizar su jerarquía de encabezados, marcar secciones para ocultar o eliminar, y exportar el documento filtrado — con soporte para selección múltiple, renumeración y filtrado desde Excel.

## Características

- **Visualización jerárquica**: muestra todos los encabezados (Heading 1–6) con nivel, numeración y sangría visual.
- **Selección múltiple**: Ctrl+Click y Shift+Click para preseleccionar filas (resaltado en celeste). La acción se aplica a todo el grupo seleccionado.
- **Acciones por encabezado**: marcar como **Ocultar** (amarillo) o **Eliminar** (rojo) individualmente o en grupo.
- **Expand / Collapse**: colapsar subárboles de encabezados para simplificar la vista.
- **Renumeración jerárquica**: reasignar numeración desde un número inicial configurable.
- **Filtrado desde Excel**: cargar un `.xlsx` y aplicar una lista de encabezados para ocultar o eliminar en lote.
- **Mostrar Oculto**: eliminar el efecto Vanish de todo el texto oculto del documento.
- **Exportación**: genera una copia del `.docx` sin las secciones marcadas, o sobreescribe el original.
- Arquitectura MVVM (.NET 8 WPF), publicación como ejecutable único autocontenido.

## Estructura del Proyecto

```
IDO_GEN_WordManager/
├── MainWindow.xaml              # UI principal: toolbar, DataGrid, barra de estado
├── MainWindow.xaml.cs           # Code-behind: selección múltiple, ícono de ventana
├── ViewModels/
│   ├── MainViewModel.cs         # MVVM: comandos, lógica de filtrado y exportación
│   └── RelayCommand.cs
├── Models/
│   └── DocumentHeading.cs       # Modelo de encabezado con estado, acción y selección
├── Services/
│   ├── WordReaderService.cs     # Lectura de encabezados desde .docx (OpenXML)
│   ├── WordExporterService.cs   # Exportación con filtrado de secciones
│   └── ExcelReaderService.cs    # Lectura de listas desde .xlsx
├── Converters/
│   ├── BoolToVisibilityConverter.cs
│   └── IndentToMarginConverter.cs
├── Styles/
│   └── AppStyles.xaml           # Paleta de colores y estilos globales
└── Resources/
    ├── logo-idom.png
    └── pen.ico
```

## Flujo de Uso

1. **Cargar Word**: selecciona un `.docx` — se listan todos los encabezados con nivel y numeración.
2. **Seleccionar encabezados**: clic simple, Ctrl+Click (múltiple) o Shift+Click (rango).
3. **Aplicar acción**: clic en **👁 Ocultar** o **🗑 Eliminar** — si hay grupo preseleccionado, la acción se aplica a todos.
4. *(Opcional)* **Filtrado por Excel**: cargar un `.xlsx`, seleccionar hoja y columna, elegir modo (Mantener / Aplicar a los de la lista) y acción (Ocultar / Eliminar), luego **Aplicar**.
5. **Exportar Word**: genera el documento sin las secciones marcadas.

## Compilar y Publicar

```bash
dotnet build -c Release
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

Resultado: `publish\IDO_GEN_WordManager.exe` (~160 MB, sin instalación requerida).

## Instalación

La aplicación se distribuye como un **ejecutable único autocontenido** (`.exe` ~160 MB). No requiere instalación de .NET ni dependencias adicionales.

1. Descargue el archivo `IDO_GEN_WordManager.rar` desde la carpeta `publish/` del repositorio
2. Descomprima y ejecute `IDO_GEN_WordManager.exe` — no requiere instalación adicional
3. En el primer arranque Windows puede mostrar SmartScreen; haga clic en **«Ejecutar de todas formas»**

## Requisitos

### Para usuarios finales

- Sistema operativo Windows 10/11
- No requiere instalación de .NET ni dependencias adicionales

### Para desarrolladores

- .NET 8 SDK
- Visual Studio 2022+ o VS Code con extensión C#
- DocumentFormat.OpenXml (incluido vía NuGet)

## Soporte Técnico

### Contacto

- **Desarrollador Principal**: Cristian Maldonado R.
- **Empresa**: IDOM Consulting, Engineering, Architecture
- **Email**: <cristian.maldonado@idom.com>

## Licencia

Copyright (c) 2026 IDOM Consulting, Engineering, Architecture. Todos los derechos reservados.

Se concede permiso, de forma gratuita, a cualquier empleado o contratista autorizado de IDOM ("Usuario Autorizado") para utilizar el Software únicamente para fines internos y comerciales de IDOM.

El Software se proporciona "tal cual", sin garantía de ningún tipo, expresa o implícita, incluidas, entre otras, las garantías de comerciabilidad, idoneidad para un propósito particular y no infracción. En ningún caso los autores o titulares de los derechos de autor serán responsables de ninguna reclamación, daño u otra responsabilidad, ya sea en una acción de contrato, agravio o de otro tipo, que surja de, fuera de o en conexión con el Software o el uso u otras relaciones en el Software.

Queda estrictamente prohibida la distribución, reproducción, sublicencia, venta, alquiler, arrendamiento o transferencia del Software, en su totalidad o en parte, a cualquier tercero sin el consentimiento previo por escrito de IDOM.

Cualquier uso no autorizado del Software dará lugar a la terminación inmediata de los derechos otorgados en virtud de esta licencia y podrá dar lugar a acciones legales.

El uso de este Software implica la aceptación de estos términos y condiciones.

![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
