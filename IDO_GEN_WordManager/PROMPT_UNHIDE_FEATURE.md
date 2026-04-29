# IDO_GEN_WordManager: Add "Unhide (Remove Vanish)" Feature

## Descripción de la funcionalidad
Agregar la capacidad de eliminar el efecto "Hidden" (Vanish) de Word de todo el texto en un documento .docx, revirtiendo el ocultamiento programático de segmentos.

## Requisitos
- Proyecto: WPF (.NET 8.0-windows)
- Biblioteca: DocumentFormat.OpenXml para manipulación de .docx
- Patrón MVVM con ViewModels observables

## Cambios requeridos

### 1. WordExporterService.cs
**Agregar método público:**
```csharp
/// <summary>
/// Elimina el efecto "Hidden" (Vanish) de todos los runs del documento,
/// haciendo visible todo el texto que fue ocultado programáticamente.
/// </summary>
public int UnhideAllVanish(string filePath)
{
	int count = 0;
	using var doc = WordprocessingDocument.Open(filePath, true);
	var body = doc.MainDocumentPart?.Document?.Body;
	if (body == null) return count;

	foreach (var run in body.Descendants<Run>())
	{
		var rpr = run.RunProperties;
		if (rpr?.Vanish != null)
		{
			rpr.Vanish.Remove();
			count++;
			// Si RunProperties quedó vacío, lo limpiamos también
			if (!rpr.HasChildren)
				rpr.Remove();
		}
	}

	doc.MainDocumentPart!.Document.Save();
	return count;
}
```

**Lógica:**
- Abre el archivo .docx en modo escritura
- Recorre todos los `Run` (segmentos de texto)
- Busca y elimina la propiedad `Vanish` de `RunProperties`
- Limpia `RunProperties` si queda vacío
- Retorna el contador de elementos restaurados

### 2. MainViewModel.cs
**Agregar propiedad:**
```csharp
public ICommand UnhideCommand { get; }
```

**Inicializar en constructor:**
```csharp
UnhideCommand = new RelayCommand(_ => UnhideDocument(), _ => IsFileLoaded);
```

**Agregar método:**
```csharp
private void UnhideDocument()
{
	try
	{
		var confirm = MessageBox.Show(
			$"¿Eliminar el efecto 'Oculto' (Vanish) de todo el texto en:\n{FilePath}?\n\nEl archivo se modificará directamente.",
			"Confirmar Mostrar Oculto", MessageBoxButton.YesNo, MessageBoxImage.Question);
		if (confirm != MessageBoxResult.Yes) return;

		int removed = _exporter.UnhideAllVanish(FilePath);
		StatusMessage = removed > 0
			? $"Se eliminó el efecto 'Oculto' de {removed} segmento(s) de texto."
			: "No se encontró texto oculto con efecto Vanish en el documento.";
		MessageBox.Show(StatusMessage, "Mostrar Oculto", MessageBoxButton.OK, MessageBoxImage.Information);
	}
	catch (Exception ex)
	{
		MessageBox.Show($"Error al eliminar el efecto oculto:\n{ex.Message}",
			"Error", MessageBoxButton.OK, MessageBoxImage.Error);
	}
}
```

**Lógica:**
- Pide confirmación del usuario con dialogo
- Llama a `UnhideAllVanish` del servicio
- Actualiza `StatusMessage` con resultado
- Maneja excepciones con dialogo de error

### 3. MainWindow.xaml
**En el Grid.ColumnDefinitions del toolbar (después de columna 9):**
```xaml
<!-- Mostrar oculto -->
<ColumnDefinition Width="Auto"/>
<ColumnDefinition Width="12"/>
```

**Agregar botón después del botón "Exportar Word" (Grid.Column="8"):**
```xaml
<Button Grid.Column="10" Content="👁  Mostrar Oculto"
		Style="{StaticResource SecondaryButton}"
		Command="{Binding UnhideCommand}"
		ToolTip="Eliminar el efecto 'Oculto' (Vanish) de todo el texto del documento"/>
```

**Ajustes de columnas posteriores:**
- Excel → Grid.Column="12"
- Etiqueta "Inicio N°" → Grid.Column="14"
- TextBox → Grid.Column="16"
- Botón Renumerar → Grid.Column="18"

## Comportamiento esperado
1. Usuario carga un documento Word con `📂 Cargar Word`
2. Botón `👁  Mostrar Oculto` aparece habilitado
3. Al hacer clic: muestra dialogo de confirmación
4. Si confirma: elimina todas las propiedades Vanish del documento
5. Muestra mensaje con cantidad de elementos restaurados

## Consideraciones técnicas
- El archivo se modifica **directamente** (sin copia)
- Reversible: se puede volver a ocultar texto con la función `ExportFiltered(..., HeadingAction.Hide)`
- Cuenta todos los `Run` con Vanish, no solo encabezados
- Limpia `RunProperties` vacíos para evitar clutter XML

## Testing manual
1. Crear .docx con texto oculto (aplicar Hide)
2. Cargar en app
3. Clic "Mostrar Oculto"
4. Confirmar
5. Verificar en Word que el texto vuelve a ser visible

## Archivos afectados
- `/Services/WordExporterService.cs` - Método UnhideAllVanish
- `/ViewModels/MainViewModel.cs` - Comando y método UnhideDocument
- `/MainWindow.xaml` - Botón y columnas de toolbar

## Status de implementación
✅ Completado - Build exitoso sin errores
