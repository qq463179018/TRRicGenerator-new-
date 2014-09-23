using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace Ric.Ui.Converter
{
    public class ExplicitPropertyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? null : GetPropertyValue(value, (string)parameter);
        }

        private static object GetPropertyValue(object target, string name)
        {
            return (
                    from type in target.GetType().GetInterfaces()
                    from prop in type.GetProperties()
                    where prop.Name == name && prop.CanRead
                    select prop.GetValue(target, new object[0])
                ).FirstOrDefault();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
