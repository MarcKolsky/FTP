using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Transfer;

namespace Amazon_S3_FTP_Program
{
    /// <summary>
    /// Interaction logic for AmazonUpload.xaml
    /// </summary>
    public partial class AmazonUpload : Window
    {
        BackgroundWorker bw = new BackgroundWorker();     //// Creates the instance for Background Worker to run.

        /// <summary>
        /// Create the empty variables to hold information that will be accessible throughout the class.
        /// </summary>
        static string reportName;
        static string clientType;
        static string clientName;
        static string fileUploadPath;
        static string fileName;
        static string Key;

        static string BucketName = "[Bucket Name]";

        public AmazonUpload()
        {
            InitializeComponent();
            
            bw.DoWork += bw_DoWork;

            //// Initiates Background Worker
            if (!bw.IsBusy)
                bw.RunWorkerAsync();

            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
        }

        /// <summary>
        /// Clears the static strings and closes this window.
        /// </summary>
        private void CleanUp()
        {
            /////////////////////////////////
            ////  Cleaup script  hidden  ////
            /////////////////////////////////

            this.Close();
        }


        #region bw_worker

        /// <summary>
        /// This is the process that actually transfers the file from the users temp directory to the Amazon S3 server.
        /// </summary>
        private void AmazonUploadProcess()
        {
            try
            {

                ////   This populates the static strings above, and makes sure the strings hold the newest values each time this window is launched.
                reportName = MainWindow.reportName;
                clientType = MainWindow.clientType.ToLower();
                clientName = MainWindow.clientName;
                fileUploadPath = MainWindow.filePath;
                fileName = Path.GetFileName(MainWindow.filePath);
                Key = clientType.ToLower() + "/" + clientName + "/" + fileName;

                string amazonS3Key = "[Amazon Key]";
                string amazonS3SecretKey = "[Amazon Secret Key]";

                TransferUtility fileTransferUtility = new TransferUtility(new AmazonS3Client(amazonS3Key, amazonS3SecretKey, Amazon.RegionEndpoint.USEast1));                 ////  Creates the instance for connecting to the Amazon S3 server.

                TransferUtilityUploadRequest fileTransferUtilityRequest = new TransferUtilityUploadRequest
                {
                    BucketName = BucketName,                                      ////  Bucket Name = ewi_public
                    FilePath = fileUploadPath,                                    ////  Path to where the zip file exists in the current users temp folder
                    StorageClass = S3StorageClass.ReducedRedundancy,              ////  Sets the storage class
                    Key = Key,                                                    ////  The location where the zip file will be uploaded to on the Amazon S3 server
                    CannedACL = S3CannedACL.PublicRead                            ////  Sets the zip file for public reading
                };

                fileTransferUtilityRequest.UploadProgressEvent += new EventHandler<UploadProgressArgs>(backgroundworker_ProgressChanged);                    ////  This reports the transfer progress to be used by the progress bar and precent complete textblock.

                fileTransferUtilityRequest.Metadata.Add("param1", "Value1");
                fileTransferUtilityRequest.Metadata.Add("param2", "Value2");
                fileTransferUtility.Upload(fileTransferUtilityRequest);                                 ////  Initiates the actual upload process to the Amazon S3 server.
            }
            catch (AmazonS3Exception s3Exception)
            {
                System.Windows.Forms.MessageBox.Show("There was an error uploading the file. \n\n" + s3Exception);
            }
        }


        /// <summary>
        ///  Populates the progress bar and percent completed textblock.
        /// </summary>
        private void backgroundworker_ProgressChanged(object sender, UploadProgressArgs e)
        {
            percentCompletedTextBlock.Dispatcher.Invoke(new Action(delegate { percentCompletedTextBlock.Text = e.PercentDone.ToString(); }));

            uploadProgressBar.Dispatcher.Invoke(new Action(delegate { uploadProgressBar.Value = e.PercentDone; }));
        }


        /// <summary>
        ///  First step in the Background Worker.
        /// </summary>
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            AmazonUploadProcess();
        }


        /// <summary>
        ///  When Background Worker completed, this forces garbage collection and runs the cleanup method.
        /// </summary>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            CleanUp();

            MainWindow createEmail = new MainWindow();
            createEmail.OutlookEmail();
        }

        #endregion


        /// <summary>
        /// This launches this window.
        /// </summary>
        public void ShowBox()
        {
            this.Show();
        }
    }
}
