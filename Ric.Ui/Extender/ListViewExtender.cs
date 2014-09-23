using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;

namespace Ric.Ui.Extender
{
    /// <summary>
    /// This class contains a few useful extenders for the ListView
    /// </summary>
    public class ListViewExtenders : DependencyObject
    {
        public static readonly DependencyProperty AutoScrollToEndProperty =
            DependencyProperty.RegisterAttached("AutoScrollToEnd", typeof (bool), typeof (ListViewExtenders),
                new UIPropertyMetadata(default(bool), OnAutoScrollToEndChanged));

        /// <summary>
        /// Returns the value of the AutoScrollToEndProperty
        /// </summary>
        /// <param name="obj">The dependency-object whichs value should be returned</param>
        /// <returns>The value of the given property</returns>
        public static bool GetAutoScrollToEnd(DependencyObject obj)
        {
            return (bool) obj.GetValue(AutoScrollToEndProperty);
        }

        /// <summary>
        /// Sets the value of the AutoScrollToEndProperty
        /// </summary>
        /// <param name="obj">The dependency-object whichs value should be set</param>
        /// <param name="value">The value which should be assigned to the AutoScrollToEndProperty</param>
        public static void SetAutoScrollToEnd(DependencyObject obj, bool value)
        {
            obj.SetValue(AutoScrollToEndProperty, value);
        }

        /// <summary>
        /// This method will be called when the AutoScrollToEnd
        /// property was changed
        /// </summary>
        /// <param name="s">The sender (the ListView)</param>
        /// <param name="e">Some additional information</param>
        public static void OnAutoScrollToEndChanged(DependencyObject s, DependencyPropertyChangedEventArgs e)
        {
            var listView = s as ListView;
            var listViewItems = listView.Items;
            var data = listViewItems.SourceCollection as INotifyCollectionChanged;

            var scrollToEndHandler = new NotifyCollectionChangedEventHandler(
                (s1, e1) =>
                {
                    if (listView.Items.Count > 0)
                    {
                        object lastItem = listView.Items[listView.Items.Count - 1];
                        listViewItems.MoveCurrentTo(lastItem);
                        listView.ScrollIntoView(lastItem);
                    }
                });

            if ((bool) e.NewValue)
                data.CollectionChanged += scrollToEndHandler;
            else
                data.CollectionChanged -= scrollToEndHandler;
        }
    }
}
