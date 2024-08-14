using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

class Progrm
{
    static async Task Main()
    {
        string baseDirectory = @"A:\8 - Website\WEBSITE ORDERS";

        if (!Directory.Exists(baseDirectory))
        {
            Directory.CreateDirectory(baseDirectory);
        }

        Application outlookApp = new Application();
        NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        Folders folders = outlookNamespace.Folders;

        foreach (MAPIFolder folder in folders)
        {
            await ProcessFolder(folder, baseDirectory);
        }

        outlookNamespace.Logoff();

        Console.WriteLine("Process completed.");
    }

    static async Task ProcessFolder(MAPIFolder folder, string baseDirectory)
    {
        List<MailItem> emailItems = new List<MailItem>();

        foreach (object item in folder.Items)
        {   
            if (item is MailItem mailItem)
            {
                emailItems.Add(mailItem);
            }
        }

        await Task.WhenAll(emailItems.Select(mailItem => Task.Run(() => ProcessMailItem(mailItem, baseDirectory))));
        await Task.WhenAll(folder.Folders.Cast<MAPIFolder>().ToList().Select(subFolder => ProcessFolder(subFolder, baseDirectory)));
    }

    private static void ProcessMailItem(MailItem mailItem, string baseDirectory)
    {
        foreach (Attachment attachment in mailItem.Attachments)
        {
            if (!attachment.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || attachment.FileName != $"order-{Regex.Match(mailItem.Subject, @"#(.+?) ").Groups[1].Value}.xml"
                || CheckFileExists(baseDirectory, attachment.FileName)) 
            {
                continue;
            }

            string sanitizedSubject = SanitizeFileName(mailItem.Subject);
            string subDirectory = mailItem.Subject.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0 ? "Test" : "New";
            string emailFolderPath = Path.Combine(Path.Combine(baseDirectory, subDirectory), sanitizedSubject);
            string filePath = Path.Combine(emailFolderPath, attachment.FileName);

            Directory.CreateDirectory(emailFolderPath);
            attachment.SaveAsFile(filePath);

            Console.WriteLine($"Downloaded: {filePath}");
        }
    }

    static string SanitizeFileName(string fileName)
    {
        string newFileName = fileName;

        string[] substringsToRemove = { "New ", "PO#:", "FW: " };

        foreach (string substring in substringsToRemove)
        {
            newFileName = newFileName.Replace(substring, string.Empty);
        }

        string invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
        string invalidReStr = $"[{invalidChars}]";
        newFileName = Regex.Replace(newFileName, invalidReStr, "_");

        newFileName = newFileName.Replace("_", " ");

        return newFileName;
    }

    static bool CheckFileExists(string directoryPath, string fileName)
    {
        string[] files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories);

        foreach (string file in files)
        {
            string currentFileName = Path.GetFileName(file);

            if (currentFileName == fileName)
            {
                return true;
            }
        }

        return false;
    }
}
