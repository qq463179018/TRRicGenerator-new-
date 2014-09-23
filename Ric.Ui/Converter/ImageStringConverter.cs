using System;
using System.Globalization;
using System.Windows.Data;

namespace Ric.Ui.Converter
{
    public class ImageStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return string.Format(@"..\Images\flags\{0}.png", value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
