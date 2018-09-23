using System.Windows;

using Slingshot.ACS.Utilities;

namespace Slingshot.ACS
{
    /// <summary>
    /// Interaction logic for Open.xaml
    /// </summary>
    public partial class Open : Window
    {
        public Open()
        {
            InitializeComponent();
        }

        private void btnOpen_Click( object sender, RoutedEventArgs e )
        {
            lblMessage.Text = string.Empty;

            var source = ImportSource.CSVFiles;
            string fileName = string.Empty;

            bool inputIsCsv = rbInputCsv.IsChecked ?? false;
            if ( !inputIsCsv )
            {
                if ( txtInput.Text != string.Empty && txtInput.Text.Contains( ".mdb" ) )
                {
                    source = ImportSource.AccessDb;
                    fileName = txtInput.Text;
                }
                else
                {
                    lblMessage.Text = "Please choose a MS Access database file.";
                    return;
                }
            }

            AcsApi.OpenConnection( source, txtInput.Text );

            if ( AcsApi.IsConnected )
            {
                if (inputIsCsv)
                {
                    Utilities.CsvToSql.CreateTables.FromFolder( txtInput.Text, AcsApi.ConnectionString );
                }
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                this.Close();
            }
            else
            {
                lblMessage.Text = $"Could not open the MS Access database file. {AcsApi.ErrorMessage}";
            }

        }

        private void Browse_Click( object sender, RoutedEventArgs e )
        {
            bool inputIsCsv = rbInputCsv.IsChecked ?? false;
            if ( inputIsCsv )
            {
                var dirDialog = new System.Windows.Forms.FolderBrowserDialog();
                var result = dirDialog.ShowDialog();

                switch ( result )
                {
                    case System.Windows.Forms.DialogResult.OK:
                        var dir = dirDialog.SelectedPath;
                        txtInput.Text = dir;
                        txtInput.ToolTip = dir;
                        break;
                    case System.Windows.Forms.DialogResult.Cancel:
                    default:
                        txtInput.Text = null;
                        txtInput.ToolTip = null;
                        break;
                }
            }
            else
            {
                var fileDialog = new System.Windows.Forms.OpenFileDialog();
                var result = fileDialog.ShowDialog();

                switch ( result )
                {
                    case System.Windows.Forms.DialogResult.OK:
                        var file = fileDialog.FileName;
                        txtInput.Text = file;
                        txtInput.ToolTip = file;
                        break;
                    case System.Windows.Forms.DialogResult.Cancel:
                    default:
                        txtInput.Text = null;
                        txtInput.ToolTip = null;
                        break;
                }
            }
        }

        private void rbInputSource_Checked( object sender, RoutedEventArgs e )
        {
            if ( rbInputCsv != null && txtInput != null && lblInput != null )
            {
                bool inputIsCsv = rbInputCsv.IsChecked ?? false;
                txtInput.Text = string.Empty;
                lblInput.Content = inputIsCsv ? "Directory" : "Filename";
            }
        }
    }
}
