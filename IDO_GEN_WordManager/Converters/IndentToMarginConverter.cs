using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace IDO_GEN_WordManager.Converters
{
    public class IndentToMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string s && double.TryParse(s, out double indent))
                return new Thickness(indent, 0, 8, 0);
            return new Thickness(0, 0, 8, 0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
