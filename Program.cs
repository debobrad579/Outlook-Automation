using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

class Program
{
    static async Task Main()
    {
        string baseDirectory = @"A:\8 - Website\WEBSITE ORDERS";

        CreateRequiredDirectories(baseDirectory);

        Application outlookApp = new Application();
        NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        Folders folders = outlookNamespace.Folders;

        await Task.WhenAll(folders.Cast<MAPIFolder>().Select(folder => ProcessFolder(folder, baseDirectory)));

        outlookNamespace.Logoff();
        Console.WriteLine("Process completed.");
    }

    private static void CreateRequiredDirectories(string baseDirectory)
    {
        if (!Directory.Exists(baseDirectory)) { Directory.CreateDirectory(baseDirectory); }
        if (!Directory.Exists(Path.Combine(baseDirectory, "New"))) { Directory.CreateDirectory(Path.Combine(baseDirectory, "New")); }
        if (!Directory.Exists(Path.Combine(baseDirectory, "Test"))) { Directory.CreateDirectory(Path.Combine(baseDirectory, "Test")); }
    }

    private static async Task ProcessFolder(MAPIFolder folder, string baseDirectory)
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
            if (!attachment.FileName.EndsWith(".xml", StringComparison.Ordinal)
                || attachment.FileName != $"order-{Regex.Match(mailItem.Subject, @"#(.+?) ").Groups[1].Value}.xml"
                || CheckFileExists(baseDirectory, attachment.FileName))
            {
                continue;
            }

            string subDirectory = mailItem.Subject.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0 ? "Test" : "New";
            string emailFolderPath = Path.Combine(Path.Combine(baseDirectory, subDirectory), FormatFolderName(mailItem.Subject));
            string filePath = Path.Combine(emailFolderPath, attachment.FileName);

            if (!Directory.Exists(emailFolderPath)) { Directory.CreateDirectory(emailFolderPath); }
            attachment.SaveAsFile(filePath);

            Console.WriteLine($"Downloaded: {filePath}");
        }
    }

    private static string FormatFolderName(string subject)
    {
        string[] substringsToRemove = { "New", "PO#:", "FW:" };

        string folderName = subject;

        foreach (string substring in substringsToRemove)
        {
            int position = folderName.IndexOf(substring);
            if (position < 0) { continue; }
            folderName = folderName.Substring(0, position) + folderName.Substring(position + substring.Length);
        }

        return Regex.Replace(Regex.Replace(folderName, $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", " "), @"\s+", " ").Trim();
    }

    private static bool CheckFileExists(string baseDirectory, string fileName)
    {
        foreach (string file in Directory.EnumerateFiles(baseDirectory, "*.*", SearchOption.AllDirectories))
        {
            if (Path.GetFileName(file) == fileName)
            {
                return true;
            }
        }

        return false;
    }
}
