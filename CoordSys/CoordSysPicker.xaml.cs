using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for CoordinateSystemPicker.xaml
    /// </summary>
    /// <remarks>Handy reference on how to style the TreeView:
    /// <a href="http://blogs.msdn.com/b/mikehillberg/archive/2009/10/30/treeview-and-hierarchicaldatatemplate-step-by-step.aspx"/>
    /// </remarks>
    public partial class CoordSysPicker : UserControl
    {
        private CoordSysPickerVM vm = new CoordSysPickerVM();
        public CoordSysPicker()
        {
            InitializeComponent();
            this.DataContext = vm;
        }

        private void TreeViewItem_BringIntoView(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = e.OriginalSource as TreeViewItem;

            var count = VisualTreeHelper.GetChildrenCount(item);
            if (0 < count)
            {
                for (int i = count - 1; i >= 0; --i)
                {
                    var childItem = VisualTreeHelper.GetChild(item, i);
                    ((FrameworkElement)childItem).BringIntoView();
                }
            }
            else
                item.BringIntoView();

            // Make sure item has focus
            if (Keyboard.FocusedElement != item)
            {
                Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() => {
                    Keyboard.Focus(item);
                    item.Focus();
                }));
            }
        }

        private void CoordSystemTreeView_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SelectedItemHelper.Content = e.NewValue;
        }
    }
}