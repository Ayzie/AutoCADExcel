using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;

namespace AutoCADExcel
{
    public partial class SaveAsPrompt : Window
    {
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public SaveAsPrompt(string Data, Microsoft.Office.Interop.Excel.Workbook workbook)
        {
            InitializeComponent();
            Output.Text = Data;
            wb = workbook;
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();
            if (filename.Text == "" && filename.Text == "filename...")
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }

            string path = dialog.SelectedPath;
            try
            {
                Globals.doc.Editor.WriteMessage("{0}{1}", path, filename.Text);
                wb.SaveAs(path + "/" + filename.Text);
            }
            catch (System.Exception ex)
            {
                Globals.doc.Editor.WriteMessage("Error: cant write file. Make sure Excel is not open. \n {0} \n", ex);
            }

            dialog.Dispose();
            this.Close();
        }

        private void filename_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            filename.Text = "";
        }

        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            wb.Close(false);
            this.Close();
        }
    }
}
