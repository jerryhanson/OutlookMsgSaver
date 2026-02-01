using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookMsgSaver
{
    public partial class MsgSaverRibbon
    {
        private void MsgSaverRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
    private void btnSaveMsg_Click(object sender, RibbonControlEventArgs e)
		{
			// 1. Get the current active Explorer and Folder
			Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
			Outlook.MAPIFolder selectedFolder = explorer.CurrentFolder;

			if (selectedFolder == null || selectedFolder.Items.Count == 0)
			{
				MessageBox.Show("Please select a folder containing emails.", "No Emails Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// 2. Open Folder Browser Dialog to let user pick destination
			using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
			{
				folderDialog.Description = $"Select a folder to save emails from '{selectedFolder.Name}'";

				if (folderDialog.ShowDialog() == DialogResult.OK)
				{
					string savePath = folderDialog.SelectedPath;
					SaveEmailsToPath(selectedFolder, savePath);
				}
			}
		}
		private void SaveEmailsToPath(Outlook.MAPIFolder folder, string directoryPath)
		{
			int successCount = 0;
			int errorCount = 0;

			// Use a loop to iterate through items
			// Note: Outlook collections are 1-based, but foreach works fine
			foreach (object item in folder.Items)
			{
				if (item is Outlook.MailItem mailItem)
				{
					try
					{
						// 3. Generate a valid filename
						string subject = mailItem.Subject;
						if (string.IsNullOrEmpty(subject)) subject = "No Subject";

						// Sanitize filename (remove illegal characters < > : " / \ | ? *)
						string cleanSubject = string.Join("_", subject.Split(Path.GetInvalidFileNameChars()));

						// Trim if too long
						if (cleanSubject.Length > 50) cleanSubject = cleanSubject.Substring(0, 50);

						// Add timestamp to ensure uniqueness
						string fileName = $"{cleanSubject}_{mailItem.ReceivedTime:yyyyMMdd_HHmmss}.msg";
						string fullPath = Path.Combine(directoryPath, fileName);

						// 4. Save the file
						mailItem.SaveAs(fullPath, Outlook.OlSaveAsType.olMSG);
						successCount++;
					}
					catch (Exception ex)
					{
						errorCount++;
						// Ideally log this error to a text file
					}
					finally
					{
						// Explicitly release COM object to prevent memory bloat in large folders
						Marshal.ReleaseComObject(mailItem);
					}
				}
			}

			MessageBox.Show($"Operation Complete.\n\nSaved: {successCount}\nErrors: {errorCount}",
							"Export Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
