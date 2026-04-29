using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace IDO_GEN_WordManager.Models
{
    public enum HeadingAction { Hide, Delete }

    public class DocumentHeading : INotifyPropertyChanged
    {
        private bool _isVisible = true;
        private bool _isExpanded = true;
        private bool _isHiddenByCollapse = false;
        private HeadingAction _action = HeadingAction.Hide;

        private string _wordNumber = string.Empty;

        // Persistent internal ID for reference
        public Guid InternalId { get; set; } = Guid.NewGuid();

        public int ParagraphIndex { get; set; }
        public int Level { get; set; }
        public string Text { get; set; } = string.Empty;

        public string WordNumber
        {
            get => _wordNumber;
            set { if (_wordNumber != value) { _wordNumber = value; OnPropertyChanged(); OnPropertyChanged(nameof(NumberLabel)); } }
        }

        public bool HasChildren { get; set; }

        public string NumberLabel => WordNumber;

        public HeadingAction Action
        {
            get => _action;
            set
            {
                if (_action != value)
                {
                    _action = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsHide));
                    OnPropertyChanged(nameof(IsDelete));
                    OnPropertyChanged(nameof(RowOpacity));
                    OnPropertyChanged(nameof(TextDecoration));
                    OnPropertyChanged(nameof(RowBackground));
                }
            }
        }

        public bool IsHide
        {
            get => !_isVisible && _action == HeadingAction.Hide;
            set
            {
                if (value)
                {
                    Action = HeadingAction.Hide;
                    IsVisible = false;
                }
                else if (_action == HeadingAction.Hide)
                {
                    IsVisible = true;
                }
            }
        }

        public bool IsDelete
        {
            get => !_isVisible && _action == HeadingAction.Delete;
            set
            {
                if (value)
                {
                    Action = HeadingAction.Delete;
                    IsVisible = false;
                }
                else if (_action == HeadingAction.Delete)
                {
                    IsVisible = true;
                }
            }
        }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    if (value) _action = HeadingAction.Hide;
                    OnPropertyChanged(nameof(IsVisible));
                    OnPropertyChanged(nameof(IsHide));
                    OnPropertyChanged(nameof(IsDelete));
                    OnPropertyChanged(nameof(RowOpacity));
                    OnPropertyChanged(nameof(TextDecoration));
                    OnPropertyChanged(nameof(RowBackground));
                }
            }
        }

        public double RowOpacity => !_isVisible && _action == HeadingAction.Delete ? 0.3 : !_isVisible ? 0.5 : 1.0;
        public string TextDecoration => !_isVisible && _action == HeadingAction.Delete ? "Strikethrough" : "None";

        /// <summary>Color de fondo de la fila: amarillo claro si oculto, rojo claro si eliminado, transparente si activo.</summary>
        public string RowBackground => !_isVisible && _action == HeadingAction.Delete ? "#FFCDD2"
                                     : !_isVisible && _action == HeadingAction.Hide    ? "#FFF9C4"
                                     : "Transparent";

        public bool IsExpanded
        {
            get => _isExpanded;
            set { if (_isExpanded != value) { _isExpanded = value; OnPropertyChanged(); OnPropertyChanged(nameof(ExpandIcon)); } }
        }

        public bool IsHiddenByCollapse
        {
            get => _isHiddenByCollapse;
            set { if (_isHiddenByCollapse != value) { _isHiddenByCollapse = value; OnPropertyChanged(); } }
        }

        public string ExpandIcon => HasChildren ? (IsExpanded ? "▼" : "▶") : string.Empty;

        public string LevelLabel => Level switch
        {
            1 => "Título",
            2 => "Subtítulo",
            3 => "Capítulo",
            _ => $"Nivel {Level}"
        };

        public string LevelIndent => ((Level - 1) * 20).ToString();

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
