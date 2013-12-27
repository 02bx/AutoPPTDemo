using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSharpeLibrary
{
    public class Dialog
    {
        public static string openDialog(string DefaultExt, string Filter)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.DefaultExt = DefaultExt;
            openDialog.Filter = Filter;
            openDialog.ShowDialog();
            return openDialog.FileName;
        }

        public static string saveDialog(string DefaultExt, string Filter)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = DefaultExt;
            saveDialog.Filter = Filter;
            saveDialog.ShowDialog();
            return saveDialog.FileName;
        }
    }
}
