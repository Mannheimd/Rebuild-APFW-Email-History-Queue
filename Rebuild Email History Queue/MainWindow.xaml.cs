using System;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Security.Principal;

namespace Rebuild_Email_History_Queue
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Check if app is running as admin.
            if (IsAdministrator() == false)
            {
                MessageBox.Show("Tool is not running as an administrator. It may be unable to stop or start APFW Outlook Service, so you will need to do this manually.");
            }

            // Work out the default history queue folder, confirm it's valid and paste it to the text box for review
            string defaultHistFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\ACT\ACT for WEB\HistoryQueue";
            if (Directory.Exists(defaultHistFolder))
            {
                textBoxFolderPath.Text = defaultHistFolder;
            }
            else
            {
                MessageBox.Show("HistoryQueue path '" + defaultHistFolder + "' does not exist or cannot be accessed. Please specify correct path.");
            }
        }

        private void buttonRebuild_Click(object sender, RoutedEventArgs e)
        {
            // Get the install path for APFW.Outlook.Srvc.exe and stop the process
            string apfwProcessPath = null;
            try
            {
                Process[] apfwProcessSearch = Process.GetProcessesByName("APFW.Outlook.Service");
                Process apfwProcess = apfwProcessSearch.First();
                apfwProcessPath = apfwProcess.MainModule.FileName;
                apfwProcess.Kill();
            }
            catch(Exception error)
            {
                MessageBox.Show("APFW.Outlook.Service.exe is not currently running or could not be stopped. Manually check it's stopped before OKing this. Rebuild will continue. Error: " + error.Message);
            }

            // Rename the HistoryQueue folder
            string originalHistFolder = textBoxFolderPath.Text;
            string newHistFolder = Directory.GetParent(originalHistFolder) + @"\HistoryQueue " + string.Format("{0:yyyy-MM-dd_hh-mm-ss}", DateTime.Now);

            try
            {
                Directory.Move(textBoxFolderPath.Text, newHistFolder);
            }
            catch(Exception error)
            {
                MessageBox.Show("Failed to rename APFW folder. Error: " + error.Message);
            }

            // Create a new HistoryQueue folder
            Directory.CreateDirectory(originalHistFolder);

            // Get a list of .msg and .xml files from the backed up folder
            string[] msgFiles = null;
            string[] xmlFiles = null;
            try
            {
                msgFiles = Directory.GetFiles(newHistFolder, "*.msg", SearchOption.TopDirectoryOnly);
            }
            catch(Exception error)
            {
                MessageBox.Show("The following error occurred whilst trying to load .msg files: " + error.Message);
            }

            try
            {
                xmlFiles = Directory.GetFiles(newHistFolder, "*.xml", SearchOption.AllDirectories);
            }
            catch (Exception error)
            {
                MessageBox.Show("The following error occurred whilst trying to load .xml files: " + error.Message);
            }

            // Loop through .msg files, copying them and their .xml files to the new HistoryQueue folder
            if (checkBoxSeparate4mb.IsChecked == true)
            {
                foreach (string msg in msgFiles)
                {
                    FileInfo msgInfo = new FileInfo(msg);
                    if (msgInfo.Length < 4000000)
                    {
                        try
                        {
                            string xml = xmlFiles.FirstOrDefault(item => item.Contains(msgInfo.Name.Replace(".msg", ".xml")));

                            string newMsgFileName = originalHistFolder + @"\" + msgInfo.Name;
                            string newXmlFileName = newMsgFileName.Replace(".msg", ".xml");

                            File.Copy(msg, newMsgFileName);
                            File.Copy(xml, newXmlFileName);
                        }
                        catch (Exception error)
                        {
                            MessageBox.Show("Copying file '" + msg + "' or it's xml failed. Error: " + error.Message);
                        }
                    }
                }
            }
            else
            {
                foreach (string msg in msgFiles)
                {
                    FileInfo msgInfo = new FileInfo(msg);
                    try
                    {
                        string xml = xmlFiles.FirstOrDefault(item => item.Contains(msgInfo.Name.Replace(".msg", ".xml")));

                        string newMsgFileName = originalHistFolder + @"\" + msgInfo.Name;
                        string newXmlFileName = newMsgFileName.Replace(".msg", ".xml");

                        File.Copy(msg, newMsgFileName);
                        File.Copy(xml, newXmlFileName);
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show("Copying file '" + msg + "' or it's xml failed. Error: " + error.Message);
                    }
                }
            }

            // Start APFW.Outlook.Service.exe - note, this will start with elevated admin rights. Can work around this by restarting this app without admin and using args to trigger the next step, but no real need.
            try
            {
                Process.Start(apfwProcessPath);
            }
            catch(Exception error)
            {
                MessageBox.Show("Failed to start APFW.Outlook.Service from '" + apfwProcessPath + "'. You may need to start this manually. Error: " + error.Message);
            }

            MessageBox.Show("Rebuild complete.");
        }

        private static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
