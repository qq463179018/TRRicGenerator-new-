using System;
using System.Windows;
using System.Windows.Controls;

namespace Ric.Ui.Control
{
    internal class ColumnDefinitionExtension : ColumnDefinition
    {
        // Variables
        public static DependencyProperty VisibleProperty;

        // Properties
        public Boolean Visible
        {
            get { return (Boolean) GetValue(VisibleProperty); }
            set { SetValue(VisibleProperty, value); }
        }

        // Constructors
        static ColumnDefinitionExtension()
        {
            VisibleProperty = DependencyProperty.Register("Visible",
                typeof (Boolean),
                typeof (ColumnDefinitionExtension),
                new PropertyMetadata(true, OnVisibleChanged));

            WidthProperty.OverrideMetadata(typeof (ColumnDefinitionExtension),
                new FrameworkPropertyMetadata(new GridLength(1, GridUnitType.Star), null,
                    CoerceWidth));

            MinWidthProperty.OverrideMetadata(typeof (ColumnDefinitionExtension),
                new FrameworkPropertyMetadata((Double) 0, null,
                    CoerceMinWidth));
        }

        // Get/Set
        public static void SetVisible(DependencyObject obj, Boolean nVisible)
        {
            obj.SetValue(VisibleProperty, nVisible);
        }

        public static Boolean GetVisible(DependencyObject obj)
        {
            return (Boolean) obj.GetValue(VisibleProperty);
        }

        private static void OnVisibleChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            obj.CoerceValue(WidthProperty);
            obj.CoerceValue(MinWidthProperty);
        }

        private static Object CoerceWidth(DependencyObject obj, Object nValue)
        {
            return (((ColumnDefinitionExtension) obj).Visible) ? nValue : new GridLength(0);
        }

        private static Object CoerceMinWidth(DependencyObject obj, Object nValue)
        {
            return (((ColumnDefinitionExtension) obj).Visible) ? nValue : (Double) 0;
        }
    }
}
