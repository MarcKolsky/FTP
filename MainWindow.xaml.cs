using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Forms;
using System.Windows.Media;

using Outlook = Microsoft.Office.Interop.Outlook;

using Amazon;
using Amazon.S3;
using Amazon.S3.Model;


namespace Amazon_S3_FTP_Program
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        BackgroundWorker bw = new BackgroundWorker();  //// Creates the instance for Background Worker to run.

        /// <summary>
        /// Create the strings to hold information that will become accessible from the AmazonUpload Class.
        /// </summary>
        public static string reportName;
        public static string clientType;
        public static string clientName;
        public static string filePath;
        static string Key;


        public MainWindow()
        {
            InitializeComponent();

            /////////////////////////////////////////
            ////  Omitted code for this sample  /////
            /////////////////////////////////////////

            ////////////////////////////////////////////////////////////////
            /// This creates the instance handler for Background Worker. ///
            /// ////////////////////////////////////////////////////////////
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            this.Closing += MainWindow_Closing;  //// This detects the closing flag for the MainWindow form, and initiates the final cleanup method below.
        }







        /// <summary>
        /// This method loops through and deletes all files that exist in the users %AppData folder where the zip files are stored.
        /// </summary>
        void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            ///////////////////////////////////////////////////////////////////////////////
            ////  Zip files stored locally are deleted here.  Code Omitted for sample  ////
            ///////////////////////////////////////////////////////////////////////////////

            Environment.Exit(0);  /// Kills the current process left open to complete this method after the MainWindow closes.
        }









        private void File_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            ///  This remains empty.  This simply detects when a dragged object has entered the Listbox.
        }


        /// <summary>
        ///  This method handles all objects dropped inside the Listbox.
        /// </summary>
        private void fileList_PreviewDrop(object sender, System.Windows.DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);  ////  This creates a string array of all items dropped inside the Listbox.  This allows the user to drop more than one file at a time.


            ///  Cycles through the above string array and adds them to the Listbox.
            if (files != null && files.Length != 0)
            {
                foreach (var file in files)
                {
                    fileList.Items.Add(file);
                }
            }
        }

        

        /// <summary>
        /// Process for removing the selected file when the "Remove File" button is clicked.
        /// </summary>
        private void RemoveFile(object sender, RoutedEventArgs e)
        {
            if (fileList.Items.Count > 0)
            {
                for (int x = fileList.SelectedItems.Count - 1; x >= 0; x--)            //  Loops through selected items in reverse
                {
                    var idx = fileList.SelectedItems[x];
                    fileList.Items.Remove(idx);
                }
            }
        }








        /// <summary>
        /// Populates the Type Combobox in the MainWindow.
        /// </summary>
        private void Type_Selected(object sender, SelectionChangedEventArgs e)
        {

            if (typeComboBox.SelectedItem.ToString() == "Clients")
            {
                ///  Pre-populate the combobox with its default value, shown when the program is launched.
                clientNameComboBox.Items.Clear();
                clientNameComboBox.SelectedIndex = 0;
                clientNameComboBox.Items.Add("------ Select Client -------");

                ///  Grabs the top folders within the Project folder on the U Drive, and adds them to the combobox.
                string[] folderNames = Directory.GetDirectories("U:\\Projects\\");
                foreach (string folder in folderNames)
                {
                    clientNameComboBox.Items.Add(folder.Substring(12));
                }
            }
            else if (typeComboBox.SelectedItem.ToString() == "Proposals")
            {
                ///  Pre-populates the combobox with its default values, shown when the user selects "Proposals" in the Type combobox
                clientNameComboBox.Items.Clear();
                clientNameComboBox.SelectedIndex = 0;
                clientNameComboBox.Items.Add("------ Select Client -------");
                clientNameComboBox.Items.Add("--- New Client ---");

                ///  This connects to the Amazon S3 server and grabs to the top "folders" within proposals.
                try
                {
                    /////////////////////////////////////////////////////////////////////////////////////
                    ////  Retrieves folders existing on the Amazon server.  Code omitted for sample  ////
                    /////////////////////////////////////////////////////////////////////////////////////
                }
                catch (AmazonS3Exception s3Exception)
                {
                    System.Windows.Forms.MessageBox.Show("There was an error. \n\n" + s3Exception);
                }
            }
        }











        /// <summary>
        ///  If the user selects the Type "proposals", and selects "New Client", this detects that change and enables the New Client Name textbox.
        /// </summary>
        private void Client_Selected(object sender, SelectionChangedEventArgs e)
        {
            if (typeComboBox.SelectedItem.ToString() == "Proposals" && clientNameComboBox.SelectedIndex == 1)
            {
                newClientNameTextBox.IsReadOnly = false;
                newClientNameTextBox.Background = Brushes.White;
            }
            else
            {
                if (newClientNameTextBox.IsEnabled == true)
                {
                    newClientNameTextBox.Clear();
                    newClientNameTextBox.IsReadOnly = true;
                    newClientNameTextBox.Background = Brushes.DarkGray;
                }
            }
        }











        ////////////////////////////////////////////////////////////
        /////////  This is where the upload process begins /////////
        ////////////////////////////////////////////////////////////

        #region File Upload

        /// <summary>
        ///  When the "Upload File(s)" button is clicked
        /// </summary>
        private void UploadFiles(object sender, RoutedEventArgs e)
        {
            ///////  Check to make sure they filled in the Report Name textbox  ///////
            if (fileNameInput.Text.Trim().Length > 0)
            {
                reportName = fileNameInput.Text;           //// Populates the static string
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please Enter Report Name!");
            }

            ///////  Check to make sure they selected a type  ///////
            if (typeComboBox.SelectedIndex > 0)
            {
                clientType = typeComboBox.SelectedItem.ToString();           //// Populates the static string
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please Select A Type!");
            }

            ///////  Check to make sure they selected a client name  ///////
            if (clientNameComboBox.SelectedIndex > 0)
            {
                if (clientNameComboBox.SelectedItem.ToString() == "--- New Client ---" && clientNameComboBox.SelectedItem.ToString() != null)       /////  If they selected a new client under proposals, this uses the text in the NEw Client Name textbox.
                {
                    if (newClientNameTextBox.Text.Trim().Length > 0)
                    {
                        clientName = newClientNameTextBox.Text;           //// Populates the static string
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Please Enter New Client Name!");
                    }
                }
                else if (clientNameComboBox.SelectedIndex > 0)
                {
                    clientName = clientNameComboBox.SelectedItem.ToString();           //// Populates the static string
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please Select A Client!");
            }

            ///////  Check to make sure they put files in the Listbox  ///////
            if (fileList.Items.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("No files have been added. \n\nPlease drag and drop your files into the box.");
            }

            //////  This is a final check to make sure everything is filled and ready to be processed.  ///////
            if (fileNameInput.Text.Trim().Length > 0 && typeComboBox.SelectedIndex > 0 && clientNameComboBox.SelectedIndex > 0 && fileList.Items.Count > 0)
            {
                if (clientNameComboBox.SelectedIndex >= 1)    /////  This is a final check to make sure they selected a client name.
                {
                    if (fileNameInput.Text.Contains("/") || fileNameInput.Text.Contains("?") || fileNameInput.Text.Contains("<") || fileNameInput.Text.Contains(">") || fileNameInput.Text.Contains("\\") || fileNameInput.Text.Contains(":") || fileNameInput.Text.Contains("*") || fileNameInput.Text.Contains("|") || fileNameInput.Text.Contains("\"") || newClientNameTextBox.Text.Contains("/") || newClientNameTextBox.Text.Contains("?") || newClientNameTextBox.Text.Contains("<") || newClientNameTextBox.Text.Contains(">") || newClientNameTextBox.Text.Contains("\\") || newClientNameTextBox.Text.Contains(":") || newClientNameTextBox.Text.Contains("*") || newClientNameTextBox.Text.Contains("|") || newClientNameTextBox.Text.Contains("\""))
                    {
                        System.Windows.Forms.MessageBox.Show("You cannot use the following characters:  /  ?  <  >  \\  :  *  |  ”");
                        return;
                    }
                    else
                    {
                        progressRing.IsActive = true;        ////  This activates the process ring that rotates in the Listbox, and is indeterminate.
                        bw.RunWorkerAsync();                 ////  This starts the Background Worker that begins the upload process.
                    }
                }
            }
            else
            {
                return;           ////  If the user fails to fill out the required information, this will return them to the MainWindow.
            }
        }











        /// <summary>
        /// This creates the zip file to be uploaded to the Amazon S3 server.
        /// </summary>
        private void ZipFiles()
        {
            if (fileList.Items.Count > 0)    //// Make sure there are actually files in the Listbox
            {
                // Specifies the location of the current users %AppData roaming folder. 
                string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

                // Combine the base folder with your specific folder....
                string specificFolder = "[Combined folder path]";

                // Check if %AppData roaming directory exists, if not, then create it
                if (!Directory.Exists(specificFolder))
                    Directory.CreateDirectory(specificFolder);

                string fileName = specificFolder + "\\" + reportName + ".zip";    ////  The location and name of the zip file

                ///  If this zip file already exists, then it will delete the file before creating the new one.  This prevents any errors.
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }


                ZipArchive zip = ZipFile.Open(fileName, ZipArchiveMode.Create);                   ////  This creates the instance for zipping the files.


                /////////////////////////////////////////////////////////////////////////////////////////////////
                ////  Loops through all files to be uploaded, and zips them here.  Code Omitted for sample.  ////
                /////////////////////////////////////////////////////////////////////////////////////////////////

                zip.Dispose();                   ////  Garbage collection.  This disposes of the zip instance.
                FileExistCheck();                ////  Now that the zip file has been created, this initiates the next method.
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No files have been added. \n\nPlease drag and drop your files into the box.");
            }

        }













        /// <summary>
        ///  This gets the top folder of the object in the Listbox.  This method keeps the ZipFiles() method from creating unneeded parent folders.
        /// </summary>
        private static string GetRightPartOfPath(string path, string startAfterPart)
        {
            // use the correct seperator for the environment
            var pathParts = path.Split(Path.DirectorySeparatorChar);

            // this assumes a case sensitive check. If you don't want this, you may want to loop through the pathParts looking
            // for your "startAfterPath" with a StringComparison.OrdinalIgnoreCase check instead
            int startAfter = Array.IndexOf(pathParts, startAfterPart);

            if (startAfter == -1)
            {
                // path path not found
                return null;
            }

            // try and work out if last part was a directory - if not, drop the last part as we don't want the filename
            var lastPartWasDirectory = pathParts[pathParts.Length - 1].EndsWith(Path.DirectorySeparatorChar.ToString());
            return string.Join(
                Path.DirectorySeparatorChar.ToString(),
                pathParts, startAfter,
                pathParts.Length - startAfter - (lastPartWasDirectory ? 0 : 1));
        }











        /// <summary>
        ///  This is the first part of the Background Worker after being initated from the Upload File(s) button.
        /// </summary>
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            ZipFiles();  ////  This initiates the method to zip the files.

            if (bw.CancellationPending == true)        ////  This processes the possibility of a cancellation of the Background Worker
            {
                e.Cancel = true;
            }
        }











        /// <summary>
        ///  When the Background Worker has completed its task of zipping the files, this method is called.
        /// </summary>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ////  If the Background Worker was canceled, then the user is returned to the MainWindow.
            if ((e.Cancelled == true))
            {
                progressRing.IsActive = false;
            }
            else
            {
                ////  The zipping is completed, and the progress ring deactivated.
                progressRing.IsActive = false;

                ////  This populates the static string for the filePath on the Amazon S3 server.
                string fileName = Path.GetFileName(MainWindow.filePath);
                Key = clientType.ToLower() + "/" + clientName + "/" + fileName;


                ////  This creates the instance to connect to the AmazonUpload Class, and initiates the method that will launch the new window.
                ////
                ////  Doing it this way keeps the program from creating more than one instance the the AmazonUpload window.
                AmazonUpload showDialog = new AmazonUpload();
                showDialog.ShowBox();

                ///   Cleanup the form for the next job.
                fileList.Items.Clear();
                fileNameInput.Clear();
                typeComboBox.SelectedIndex = 0;
                clientNameComboBox.Items.Clear();
                newClientNameTextBox.Clear();
                
            }
        }









        /// <summary>
        ///  This method verifies whether or not the file to be uploaded already exists on the Amazon S3 server.
        /// </summary>
        private void FileExistCheck()
        {
            ////  Connection credentials for the Amazon connection.
            string keyName = "[Amazon Key]";
            string secretKey = "[Amazon Secret Key]";

            AmazonS3Config config = new AmazonS3Config();
            config.ServiceURL = "[Amazon folder location]";
            config.RegionEndpoint = Amazon.RegionEndpoint.USEast1;

            AmazonS3Client proposalBucketConnection = new AmazonS3Client(keyName, secretKey, config);

            ListObjectsRequest request = new ListObjectsRequest
            {
                BucketName = "[Bucket Name]",
                Marker = clientType.ToLower() + "/"        ////  The Marker is detemined by which type was selected by the user.  It is also converted to be all lowercase.
            };


            string uploadFileName = Path.GetFileName(filePath);                                             ////  Gets the name of the file to be uploaded.
            string fileNameCheck = clientType.ToLower() + "/" + clientName + "/" + uploadFileName;          ////  This string represents the location to be checked on the server.

            ListObjectsResponse response = proposalBucketConnection.ListObjects(request);
            foreach (S3Object item in response.S3Objects)
            {
                ////  If the file already exists, the user is given the option to overwrite or rename the file they are trying to upload.
                if (item.Key == fileNameCheck)
                {
                    DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("File Already Exists.  Would you like to replace this file?", "File Already Exists", MessageBoxButtons.YesNo);
                    if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        continue;             ////  This overwrites the file currently on the Amazon S3 server.
                    }
                    else if (dialogResult == System.Windows.Forms.DialogResult.No)
                    {
                        // Cancels the asynchronous operation, and returns to the user to the MainWindow to rename the zip file.
                        bw.CancelAsync();
                    }
                }
                else
                {
                    continue;              ////  If there is no file conflict, the process continues as normal.
                }
            }
            proposalBucketConnection.Dispose();             ////  This disposes of the Amazon S3 connection instance.
        }

        #endregion


        









        /// <summary>
        ///  This generates the e-mail with the link to the files the user will send to the client.
        /// </summary>
        public void OutlookEmail()
        {
            string link = "[Link uri]";


            ///  Check to see if Outlook is already running
            Process[] Processes = Process.GetProcessesByName("OUTLOOK");
            if (Processes.Length > 0)
            {
                //This creates a new e-mail and opens it after the file upload has completed
                Outlook.Application otApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)
                    otApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.To = "";
                mailItem.HTMLBody += "[Enter text to client here!]<br>";
                mailItem.HTMLBody += "<p><a href=" + link + "> Download File Here </a><br>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += ReadSignature();
                mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
                mailItem.Display(false);
            }
            else
            {

                // If Outlook is not running then we start it
                Process.Start("Outlook.exe");

                //This creates a new e-mail and opens it after the file upload has completed
                Outlook.Application otApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)
                    otApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.To = "";
                mailItem.HTMLBody += "[Enter text to client here!]<br>";
                mailItem.HTMLBody += "<p><a href=" + link + "> Download File Here </a><br>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += "<p>";
                mailItem.HTMLBody += ReadSignature();
                mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
                mailItem.Display(false);
            }

            Key = null;                      ////  Set the static string to null.
        }









        /// Set the user's signature when creating a new email ///
        private string ReadSignature()
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);
            if
            (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                if (fiSignature.Length > 0)
                {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, System.Text.Encoding.Default);
                    signature = sr.ReadToEnd();
                    if (!string.IsNullOrEmpty(signature))
                    {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }

            }
            return signature;
        }







        private void fileNameInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (fileNameInput.Text.Contains("/") || fileNameInput.Text.Contains("?") || fileNameInput.Text.Contains("<") || fileNameInput.Text.Contains(">") || fileNameInput.Text.Contains("\\") || fileNameInput.Text.Contains(":") || fileNameInput.Text.Contains("*") || fileNameInput.Text.Contains("|") || fileNameInput.Text.Contains("\""))
            {
                e.Handled = true;
                System.Windows.Forms.MessageBox.Show("You cannot use the following characters:  /  ?  <  >  \\  :  *  |  ”");
                fileNameInput.Text = fileNameInput.Text.Substring(0, fileNameInput.Text.Length - 1);
                fileNameInput.SelectionStart = fileNameInput.Text.Length;
            }
        }







        private void newClientNameTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (newClientNameTextBox.Text.Contains("/") || newClientNameTextBox.Text.Contains("?") || newClientNameTextBox.Text.Contains("<") || newClientNameTextBox.Text.Contains(">") || newClientNameTextBox.Text.Contains("\\") || newClientNameTextBox.Text.Contains(":") || newClientNameTextBox.Text.Contains("*") || newClientNameTextBox.Text.Contains("|") || newClientNameTextBox.Text.Contains("\""))
            {
                e.Handled = true;
                System.Windows.Forms.MessageBox.Show("You cannot use the following characters:  /  ?  <  >  \\  :  *  |  ”");
                newClientNameTextBox.Text = newClientNameTextBox.Text.Substring(0, newClientNameTextBox.Text.Length - 1);
                newClientNameTextBox.SelectionStart = newClientNameTextBox.Text.Length;
            }
        }
    }
}
