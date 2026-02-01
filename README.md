# OutlookMsgSaver
This is a outlook desktop plugin which will mannualy save all emails in selected folder to .msg file. User interface is provided to set the saved folder destination.

**VSTO (Visual Studio Tools for Office) Add-in.**

Why VSTO? Modern "Web Add-ins" (Office.js) run in a sandbox and generally cannot access the local file system directly to save files without user intervention for every single file. VSTO add-ins (C#/.NET) have full access to the Windows file system and the Outlook Object Model, allowing you to save .msg files programmatically.

Here is a step-by-step guide and the complete source code to create this plugin.

## Prerequisites
OS: Windows (VSTO does not run on Mac).

IDE: Visual Studio 2022 (or 2019).

Workload: During Visual Studio installation, ensure "Office/SharePoint development" is checked.

* Step 1: Create the Project
Open Visual Studio.

Click Create a new project.

Search for Outlook VSTO Add-in.

Select the C# version and click Next.

Name it OutlookMsgSaver and click Create.

* Step 2: Add the User Interface (Ribbon)
In the Solution Explorer, right-click the project name (OutlookMsgSaver).

Select Add -> New Item.

Search for Ribbon (Visual Designer). Name it MsgSaverRibbon.cs.

Click Add.

The Ribbon Designer will appear.

Click the Group1 area on the ribbon. Change the Label property (bottom right) to "Email Saver".

Drag a Button from the Toolbox onto the group.

Change the Button Label to "Save Folder to .MSG".

Change the Button Name to btnSaveMsg.

(Optional) You can set a large icon using the OfficeImageId property (e.g., set it to FileSave).

* Step 3: The Code
Double-click the button you just created in the Designer. This will open MsgSaverRibbon.cs. Replace the contents with the following code.

You will need to add a reference to System.Windows.Forms if it is missing (Right-click Project -> Add Reference -> Assemblies -> Framework -> System.Windows.Forms).


Key Logic Breakdown
Selection: We access Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder to see what folder the user is currently looking at in Outlook.

System.Windows.Forms: We use FolderBrowserDialog to give the user a native Windows interface to pick the destination directory.

Sanitization: Email subjects often contain characters that are illegal in Windows filenames (like :, /, ?). The code replaces these with underscores and adds a timestamp to prevent overwriting files that have the same subject.

SaveAs: We use the native mailItem.SaveAs(..., olMSG) method. This preserves the email exactly as it is in Outlook, including attachments.

Memory Management: COM objects (Office Interop) can sometimes cause memory leaks if not released. The Marshal.ReleaseComObject in the finally block ensures we clean up after every email.

* Step 4: Run and Test
Press F5 in Visual Studio.

This will launch a new instance of Outlook.

Go to the Add-ins tab (or the tab name if you customized the Ribbon XML, usually it appears in a generic "Add-ins" tab or its own tab if you configured the Ribbon Designer properties).

Navigate to a folder with emails (e.g., Inbox).

Click your "Save Folder to .MSG" button.

Select a folder on your Desktop.

Check the folder to see your .msg files appear.

## How to Deploy
To give this to another user:

In Visual Studio, right-click the project -> Publish.

Follow the wizard to create an installer (setup.exe).

The user can run that installer to add the plugin to their Outlook.
