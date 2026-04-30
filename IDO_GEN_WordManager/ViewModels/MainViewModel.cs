using IDO_GEN_WordManager.Models;
using IDO_GEN_WordManager.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace IDO_GEN_WordManager.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private static readonly Regex HierarchyTokenRegex = new(@"^\d+(?:[.\-\/]\d+)+\.?$", RegexOptions.Compiled);

        private readonly WordReaderService _reader = new();
        private readonly WordExporterService _exporter = new();
        private readonly ExcelReaderService _excelReader = new();

        private string _filePath = string.Empty;
        private string _statusMessage = "Sin archivo cargado";
        private bool _isFileLoaded;
        private bool _overwriteSource = false;

        // Excel import state
        private List<ExcelSheetData> _excelWorkbook = new();
        private Dictionary<int, (string Header, List<string> Values)> _excelData = new();
        private string _excelFileName = "Ningún archivo cargado";
        private bool _isExcelPanelVisible = true;
        private string? _selectedExcelSheet;
        private string? _selectedExcelColumn;
        private int _excelModeIndex = 1;   // 0 = Mantener, 1 = Aplicar a los de la lista
        private int _excelActionIndex; // 0 = Ocultar,  1 = Eliminar

        public ObservableCollection<string> ExcelSheets { get; } = new();
        public ObservableCollection<string> ExcelColumns { get; } = new();

        public string ExcelFileName
        {
            get => _excelFileName;
            set { _excelFileName = value; OnPropertyChanged(); }
        }

        public bool IsExcelPanelVisible
        {
            get => _isExcelPanelVisible;
            set
            {
                _isExcelPanelVisible = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ExcelPanelToggleLabel));
            }
        }

        public string ExcelPanelToggleLabel => IsExcelPanelVisible ? "📊  Ocultar Excel" : "📊  Mostrar Excel";

        public string? SelectedExcelSheet
        {
            get => _selectedExcelSheet;
            set
            {
                if (_selectedExcelSheet == value) return;
                _selectedExcelSheet = value;
                OnPropertyChanged();
                LoadSelectedExcelSheetColumns();
            }
        }

        public string? SelectedExcelColumn
        {
            get => _selectedExcelColumn;
            set { _selectedExcelColumn = value; OnPropertyChanged(); }
        }

        public int ExcelModeIndex
        {
            get => _excelModeIndex;
            set { _excelModeIndex = value; OnPropertyChanged(); }
        }

        public int ExcelActionIndex
        {
            get => _excelActionIndex;
            set { _excelActionIndex = value; OnPropertyChanged(); }
        }

        private int _renumberStart = 1;
        public int RenumberStart
        {
            get => _renumberStart;
            set { _renumberStart = value < 0 ? 0 : value; OnPropertyChanged(); }
        }

        // Lista interna completa (incluye items colapsados)
        private readonly ObservableCollection<DocumentHeading> _allHeadings = new();

        // Lista visible en el DataGrid (excluye items ocultos por colapso)
        public ObservableCollection<DocumentHeading> Headings { get; } = new();

        public string FilePath
        {
            get => _filePath;
            set { _filePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(FileName)); }
        }

        public string FileName => string.IsNullOrEmpty(_filePath) ? "—" : Path.GetFileName(_filePath);

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public bool IsFileLoaded
        {
            get => _isFileLoaded;
            set { _isFileLoaded = value; OnPropertyChanged(); }
        }

        public bool OverwriteSource
        {
            get => _overwriteSource;
            set
            {
                if (_overwriteSource == value) return;
                _overwriteSource = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CreateCopy));
                OnPropertyChanged(nameof(OverwriteOriginal));
                OnPropertyChanged(nameof(ExportModeLabel));
            }
        }

        public bool CreateCopy
        {
            get => !_overwriteSource;
            set
            {
                if (value) OverwriteSource = false;
            }
        }

        public bool OverwriteOriginal
        {
            get => _overwriteSource;
            set
            {
                if (value) OverwriteSource = true;
            }
        }

        public string ExportModeLabel => _overwriteSource ? "Sobreescribir archivo original" : "Crear copia";

        public int VisibleCount => _allHeadings.Count(h => h.IsVisible);
        public int TotalCount => _allHeadings.Count;

        public ICommand LoadFileCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }
        public ICommand ToggleExpandCommand { get; }
        public ICommand ToggleExcelPanelCommand { get; }
        public ICommand LoadExcelCommand { get; }
        public ICommand ApplyExcelCommand { get; }
        public ICommand RenumberCommand { get; }
        public ICommand UnhideCommand { get; }

        public MainViewModel()
        {
            LoadFileCommand         = new RelayCommand(_ => LoadFile());
            ExportCommand           = new RelayCommand(_ => Export(), _ => IsFileLoaded && _allHeadings.Count > 0);
            SelectAllCommand        = new RelayCommand(_ => SetAllVisibility(true),  _ => _allHeadings.Count > 0);
            DeselectAllCommand      = new RelayCommand(_ => SetAllVisibility(false), _ => _allHeadings.Count > 0);
            ToggleExpandCommand     = new RelayCommand(h => ToggleExpand(h as DocumentHeading));
            ToggleExcelPanelCommand = new RelayCommand(_ => IsExcelPanelVisible = !IsExcelPanelVisible);
            LoadExcelCommand        = new RelayCommand(_ => LoadExcel());
            ApplyExcelCommand       = new RelayCommand(_ => ApplyExcel(),
                                        _ => IsFileLoaded && SelectedExcelColumn != null);
            RenumberCommand         = new RelayCommand(_ => Renumber(),
                                        _ => IsFileLoaded && _allHeadings.Count > 0);
            UnhideCommand           = new RelayCommand(_ => UnhideDocument(), _ => IsFileLoaded);
        }

        private void LoadFile()
        {
            var dlg = new OpenFileDialog
            {
                Title = "Seleccionar documento Word",
                Filter = "Documentos Word (*.docx)|*.docx"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                FilePath = dlg.FileName;
                var headings = _reader.LoadHeadings(FilePath);

                foreach (var old in _allHeadings) old.PropertyChanged -= Heading_PropertyChanged;
                _allHeadings.Clear();
                Headings.Clear();

                foreach (var h in headings)
                {
                    h.PropertyChanged += Heading_PropertyChanged;
                    _allHeadings.Add(h);
                    Headings.Add(h);
                }

                IsFileLoaded = true;
                RefreshCounters();
                StatusMessage = _allHeadings.Count == 0
                    ? "No se encontraron encabezados (Heading 1-3) en el documento."
                    : $"Se cargaron {_allHeadings.Count} encabezados correctamente.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error al cargar el archivo: {ex.Message}";
                MessageBox.Show($"No se pudo cargar el documento:\n\n{ex.Message}",
                    "Error al cargar", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

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

        private void Renumber()
        {
            var counters = new int[10];
            counters[1] = RenumberStart - 1;

            foreach (var h in _allHeadings)
            {
                var lvl = Math.Clamp(h.Level, 1, 9);
                counters[lvl]++;
                for (int i = lvl + 1; i < counters.Length; i++) counters[i] = 0;

                var parts = new List<string>();
                for (int i = 1; i <= lvl; i++) parts.Add(counters[i].ToString());
                h.WordNumber = string.Join(".", parts);
            }

            StatusMessage = $"Numeración reasignada desde {RenumberStart}.";
        }

        private void LoadExcel()
        {
            var dlg = new OpenFileDialog
            {
                Title = "Seleccionar archivo Excel",
                Filter = "Archivos Excel (*.xlsx)|*.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                _excelWorkbook = _excelReader.LoadWorkbook(dlg.FileName);
                ExcelFileName = Path.GetFileName(dlg.FileName);

                ExcelSheets.Clear();
                foreach (var sheet in _excelWorkbook)
                    ExcelSheets.Add(sheet.SheetName);

                ExcelColumns.Clear();
                SelectedExcelSheet = ExcelSheets.FirstOrDefault();
                StatusMessage = $"Excel cargado: {ExcelSheets.Count} hoja(s) detectadas.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error al cargar Excel: {ex.Message}";
                MessageBox.Show($"No se pudo cargar el archivo Excel:\n\n{ex.Message}",
                    "Error al cargar Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyExcel()
        {
            if (SelectedExcelColumn == null || _excelData.Count == 0) return;

            var colEntry = _excelData.Values.FirstOrDefault(c => c.Header == SelectedExcelColumn);
            var selectedNumbers = new HashSet<string>(
                (colEntry.Values ?? new List<string>())
                    .Select(NormalizeHierarchyKey)
                    .Where(v => !string.IsNullOrWhiteSpace(v))!
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            var availableNumbers = new HashSet<string>(
                _allHeadings
                    .Select(h => NormalizeHierarchyKey(h.WordNumber))
                    .Where(v => !string.IsNullOrWhiteSpace(v))!,
                StringComparer.OrdinalIgnoreCase);

            bool keepMode    = ExcelModeIndex == 0;   // 0=Mantener, 1=Aplicar
            bool deleteAction = ExcelActionIndex == 1; // 0=Ocultar,  1=Eliminar

            int affected = 0;
            foreach (var h in _allHeadings)
            {
                var num = NormalizeHierarchyKey(h.WordNumber);
                bool match = !string.IsNullOrEmpty(num) && selectedNumbers.Contains(num);

                bool shouldMark = keepMode ? !match : match;
                if (shouldMark)
                {
                    if (deleteAction) h.IsDelete = true;
                    else              h.IsHide   = true;
                    affected++;
                }
                else
                {
                    h.IsVisible = true;
                }
            }

            RefreshCounters();
            var actionLabel = deleteAction ? "eliminar" : "ocultar";
            var missingNumbers = selectedNumbers.Where(n => !availableNumbers.Contains(n)).ToList();
            if (affected == 0 && missingNumbers.Count > 0)
            {
                var preview = string.Join(", ", missingNumbers.Take(5));
                StatusMessage = $"No se encontraron referencias en la numeracion visible. Excel no hallado: {preview}";
            }
            else if (missingNumbers.Count > 0)
            {
                var preview = string.Join(", ", missingNumbers.Take(5));
                StatusMessage = $"Excel aplicado sobre la numeracion visible: {affected} encabezado(s) marcados para {actionLabel}. No encontrados: {preview}";
            }
            else
            {
                StatusMessage = $"Excel aplicado sobre la numeracion visible de la interfaz: {affected} encabezado(s) marcados para {actionLabel}.";
            }
        }

        private static string NormalizeHierarchyKey(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = value.Trim()
                .Replace('\u00A0', ' ')
                .Trim()
                .Trim('\'', '"');

            if (!HierarchyTokenRegex.IsMatch(normalized))
                return normalized;

            normalized = normalized.TrimEnd('.').Replace('-', '.').Replace('/', '.');
            var parts = normalized.Split('.', StringSplitOptions.RemoveEmptyEntries);

            return string.Join(".", parts.Select(p =>
                int.TryParse(p, out var n) ? n.ToString() : p.TrimStart('0')));
        }

        private void LoadSelectedExcelSheetColumns()
        {
            ExcelColumns.Clear();
            _excelData = new();

            if (string.IsNullOrWhiteSpace(SelectedExcelSheet))
            {
                SelectedExcelColumn = null;
                return;
            }

            var sheet = _excelWorkbook.FirstOrDefault(s => string.Equals(s.SheetName, SelectedExcelSheet, StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                SelectedExcelColumn = null;
                return;
            }

            _excelData = sheet.Columns;
            foreach (var kv in _excelData.OrderBy(k => k.Key))
                ExcelColumns.Add(kv.Value.Header);

            SelectedExcelColumn = ExcelColumns.FirstOrDefault();
        }

        private void ToggleExpand(DocumentHeading? parent)
        {
            if (parent == null || !parent.HasChildren) return;

            parent.IsExpanded = !parent.IsExpanded;

            // Encontrar todos los hijos directos e indirectos
            var parentIdx = _allHeadings.IndexOf(parent);
            for (int i = parentIdx + 1; i < _allHeadings.Count; i++)
            {
                var child = _allHeadings[i];
                if (child.Level <= parent.Level) break;

                if (!parent.IsExpanded)
                {
                    child.IsHiddenByCollapse = true;
                    Headings.Remove(child);
                }
                else
                {
                    // Solo mostrar si el padre inmediato está expandido
                    var immediateParent = FindImmediateParent(i);
                    if (immediateParent == null || immediateParent.IsExpanded)
                    {
                        child.IsHiddenByCollapse = false;
                        var insertAt = FindInsertPosition(child);
                        Headings.Insert(insertAt, child);
                    }
                }
            }
        }

        private DocumentHeading? FindImmediateParent(int childIdx)
        {
            var child = _allHeadings[childIdx];
            for (int i = childIdx - 1; i >= 0; i--)
            {
                if (_allHeadings[i].Level < child.Level)
                    return _allHeadings[i];
            }
            return null;
        }

        private int FindInsertPosition(DocumentHeading item)
        {
            var allIdx = _allHeadings.IndexOf(item);
            // Buscar el primer item que en _allHeadings viene después y ya está en Headings
            for (int i = allIdx + 1; i < _allHeadings.Count; i++)
            {
                var next = _allHeadings[i];
                var visIdx = Headings.IndexOf(next);
                if (visIdx >= 0) return visIdx;
            }
            return Headings.Count;
        }

        private void Export()
        {
            try
            {
                string destPath;
                if (OverwriteSource)
                {
                    var confirm = MessageBox.Show(
                        $"¿Sobreescribir el archivo original?\n{FilePath}",
                        "Confirmar sobreescritura", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (confirm != MessageBoxResult.Yes) return;
                    destPath = FilePath;
                }
                else
                {
                    var dlg = new SaveFileDialog
                    {
                        Title = "Guardar copia del documento",
                        Filter = "Documentos Word (*.docx)|*.docx",
                        FileName = Path.GetFileNameWithoutExtension(FilePath) + "_filtrado.docx",
                        InitialDirectory = Path.GetDirectoryName(FilePath)
                    };
                    if (dlg.ShowDialog() != true) return;
                    destPath = dlg.FileName;
                }

                _exporter.ExportFiltered(FilePath, destPath, _allHeadings);
                StatusMessage = $"Exportado: {Path.GetFileName(destPath)}";
                MessageBox.Show($"Documento guardado correctamente:\n{destPath}",
                    "Exportación exitosa", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SetAllVisibility(bool visible)
        {
            foreach (var h in _allHeadings) h.IsVisible = visible;
            RefreshCounters();
        }

        public event EventHandler? ClearSelectionRequested;

        private bool _applyingGroupAction = false;

        private void Heading_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(DocumentHeading.IsVisible))
                RefreshCounters();

            if (_applyingGroupAction) return;
            if (e.PropertyName != nameof(DocumentHeading.ActionApplied)) return;

            var changed = sender as DocumentHeading;
            if (changed == null || !changed.IsSelected) return;

            var selected = Headings.Where(h => h.IsSelected && h != changed).ToList();
            if (selected.Count == 0) return;

            _applyingGroupAction = true;
            try
            {
                foreach (var h in selected)
                {
                    h.Action     = changed.Action;
                    h.IsVisible  = changed.IsVisible;
                    h.IsSelected = false;
                }
                changed.IsSelected = false;
                RefreshCounters();
                ClearSelectionRequested?.Invoke(this, EventArgs.Empty);
            }
            finally
            {
                _applyingGroupAction = false;
            }
        }

        private void RefreshCounters()
        {
            OnPropertyChanged(nameof(VisibleCount));
            OnPropertyChanged(nameof(TotalCount));
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
