// Builds the Windows Installer package straight through msi.dll, so no WiX or other
// toolset is needed: everything used here ships with Windows.
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;

static class MsiBuilder
{
    const string UpgradeCode = "{7E3B6A2C-4D1F-4B8E-9C5A-2F6D8E1B3A70}";
    const string ComponentId = "{B2C4E6A8-1357-4A9B-8C0D-E2F4A6B8C0D1}";
    const string ProductName = "Outlook Notifier";
    const string Manufacturer = "Outlook Notifier";
    const string ExeName = "OutlookNotifier.exe";

    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiOpenDatabase(string path, IntPtr persist, out IntPtr db);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiDatabaseOpenView(IntPtr db, string query, out IntPtr view);
    [DllImport("msi.dll")] static extern uint MsiViewExecute(IntPtr view, IntPtr record);
    [DllImport("msi.dll")] static extern uint MsiViewClose(IntPtr view);
    [DllImport("msi.dll")] static extern uint MsiCloseHandle(IntPtr handle);
    [DllImport("msi.dll")] static extern IntPtr MsiCreateRecord(uint fields);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiRecordSetString(IntPtr record, uint field, string value);
    [DllImport("msi.dll")] static extern uint MsiRecordSetInteger(IntPtr record, uint field, int value);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiRecordSetStream(IntPtr record, uint field, string path);
    [DllImport("msi.dll")] static extern uint MsiDatabaseCommit(IntPtr db);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiGetSummaryInformation(IntPtr db, string path, uint updateCount, out IntPtr summary);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiSummaryInfoSetProperty(IntPtr summary, uint property, uint dataType, int value, ref FILETIME time, string text);
    [DllImport("msi.dll")] static extern uint MsiSummaryInfoPersist(IntPtr summary);
    [DllImport("msi.dll", CharSet = CharSet.Unicode)] static extern uint MsiGetFileVersion(string path, StringBuilder version, ref uint versionLength, StringBuilder lang, ref uint langLength);

    class FileData { public string Path; public FileData(string p) { Path = p; } }

    static IntPtr db;

    static int Main(string[] args)
    {
        if (args.Length != 4)
        {
            Console.Error.WriteLine("usage: MsiBuilder <app.exe> <app.ico> <output.msi> <workdir>");
            return 2;
        }
        try { Build(Path.GetFullPath(args[0]), Path.GetFullPath(args[1]), Path.GetFullPath(args[2]), Path.GetFullPath(args[3])); return 0; }
        catch (Exception e) { Console.Error.WriteLine("error: " + e.Message); return 1; }
    }

    static void Build(string exe, string icon, string msi, string work)
    {
        string fileVersion = GetFileVersion(exe);
        string productVersion = string.Join(".", fileVersion.Split('.').Take(3));
        string cab = MakeCabinet(exe, work);
        if (File.Exists(msi)) File.Delete(msi);

        Check(MsiOpenDatabase(msi, (IntPtr)3 /* MSIDBOPEN_CREATE */, out db), "MsiOpenDatabase");
        try
        {
            CreateTables();

            string productCode = NewGuid();
            Insert("Property", "Property,Value", "ProductCode", productCode);
            Insert("Property", "Property,Value", "ProductName", ProductName);
            Insert("Property", "Property,Value", "ProductVersion", productVersion);
            Insert("Property", "Property,Value", "ProductLanguage", "1033");
            Insert("Property", "Property,Value", "Manufacturer", Manufacturer);
            Insert("Property", "Property,Value", "UpgradeCode", UpgradeCode);
            Insert("Property", "Property,Value", "ALLUSERS", "1");
            Insert("Property", "Property,Value", "ARPPRODUCTICON", "AppIcon.ico");
            Insert("Property", "Property,Value", "ARPNOMODIFY", "1");
            Insert("Property", "Property,Value", "ARPNOREPAIR", "1");
            Insert("Property", "Property,Value", "SecureCustomProperties", "PREVIOUSVERSIONSINSTALLED");

            Insert("Directory", "Directory,Directory_Parent,DefaultDir", "TARGETDIR", null, "SourceDir");
            Insert("Directory", "Directory,Directory_Parent,DefaultDir", "ProgramFiles64Folder", "TARGETDIR", ".");
            Insert("Directory", "Directory,Directory_Parent,DefaultDir", "INSTALLFOLDER", "ProgramFiles64Folder", "OUTLOO~1|Outlook Notifier");
            Insert("Directory", "Directory,Directory_Parent,DefaultDir", "ProgramMenuFolder", "TARGETDIR", ".");

            Insert("Component", "Component,ComponentId,Directory_,Attributes,Condition,KeyPath",
                "MainComponent", ComponentId, "INSTALLFOLDER", 256 /* 64-bit */, null, "AppExe");
            Insert("File", "File,Component_,FileName,FileSize,Version,Language,Attributes,Sequence",
                "AppExe", "MainComponent", "OUTLOO~1.EXE|" + ExeName, (int)new FileInfo(exe).Length, fileVersion, "0", 0, 1);
            Insert("Media", "DiskId,LastSequence,DiskPrompt,Cabinet,VolumeLabel,Source", 1, 1, null, "#app.cab", null, null);
            Insert("_Streams", "Name,Data", "app.cab", new FileData(cab));

            Insert("Feature", "Feature,Feature_Parent,Title,Description,Display,Level,Directory_,Attributes",
                "MainFeature", null, ProductName, "Notifications des nouveaux e-mails Outlook", 1, 1, "INSTALLFOLDER", 0);
            Insert("FeatureComponents", "Feature_,Component_", "MainFeature", "MainComponent");

            // Start with Windows for every user; each user can turn it off in the app settings.
            Insert("Registry", "Registry,Root,Key,Name,Value,Component_",
                "AutoStart", 2, @"Software\Microsoft\Windows\CurrentVersion\Run", "OutlookNotifier", "\"[#AppExe]\" --autostart", "MainComponent");

            Insert("Shortcut", "Shortcut,Directory_,Name,Component_,Target,Arguments,Description,Hotkey,Icon_,IconIndex,ShowCmd,WkDir",
                "StartMenu", "ProgramMenuFolder", "OUTLOO~1|Outlook Notifier", "MainComponent", "[#AppExe]", null,
                "Notifications des nouveaux e-mails Outlook", null, null, null, null, "INSTALLFOLDER");

            Insert("Icon", "Name,Data", "AppIcon.ico", new FileData(icon));

            // Major upgrade: any earlier or identical version is removed first.
            Insert("Upgrade", "UpgradeCode,VersionMin,VersionMax,Language,Attributes,Remove,ActionProperty",
                UpgradeCode, "0.0.0", productVersion, null, 256 | 512 | 1, null, "PREVIOUSVERSIONSINSTALLED");

            // 34 = run exe in a directory, 64 = ignore exit code.
            Insert("CustomAction", "Action,Type,Source,Target",
                "StopRunningApp", 34 + 64, "INSTALLFOLDER", "\"[INSTALLFOLDER]" + ExeName + "\" --quit");
            // Launch through the shell so the app runs as the signed-in user, never elevated
            // (an elevated process cannot reach a normal Outlook over COM).
            Insert("CustomAction", "Action,Type,Source,Target",
                "LaunchApp", 34 + 64, "INSTALLFOLDER", "\"[WindowsFolder]explorer.exe\" \"[#AppExe]\"");

            object[][] execute =
            {
                new object[] { "FindRelatedProducts", null, 25 },
                new object[] { "CostInitialize", null, 800 },
                new object[] { "FileCost", null, 900 },
                new object[] { "CostFinalize", null, 1000 },
                new object[] { "StopRunningApp", "REMOVE~=\"ALL\" OR PREVIOUSVERSIONSINSTALLED", 1390 },
                new object[] { "InstallValidate", null, 1400 },
                new object[] { "RemoveExistingProducts", null, 1401 },
                new object[] { "InstallInitialize", null, 1500 },
                new object[] { "ProcessComponents", null, 1600 },
                new object[] { "UnpublishFeatures", null, 1800 },
                new object[] { "RemoveRegistryValues", null, 2600 },
                new object[] { "RemoveShortcuts", null, 3200 },
                new object[] { "RemoveFiles", null, 3500 },
                new object[] { "InstallFiles", null, 4000 },
                new object[] { "CreateShortcuts", null, 4500 },
                new object[] { "WriteRegistryValues", null, 5000 },
                new object[] { "RegisterUser", null, 6000 },
                new object[] { "RegisterProduct", null, 6100 },
                new object[] { "PublishFeatures", null, 6300 },
                new object[] { "PublishProduct", null, 6400 },
                new object[] { "InstallFinalize", null, 6600 },
                new object[] { "LaunchApp", "NOT REMOVE", 6700 }
            };
            foreach (var row in execute) Insert("InstallExecuteSequence", "Action,Condition,Sequence", row);

            object[][] ui =
            {
                new object[] { "FindRelatedProducts", null, 25 },
                new object[] { "CostInitialize", null, 800 },
                new object[] { "FileCost", null, 900 },
                new object[] { "CostFinalize", null, 1000 },
                new object[] { "ExecuteAction", null, 1300 }
            };
            foreach (var row in ui) Insert("InstallUISequence", "Action,Condition,Sequence", row);

            WriteSummary();
            Check(MsiDatabaseCommit(db), "MsiDatabaseCommit");
        }
        finally { MsiCloseHandle(db); }

        Console.WriteLine("MSI: " + msi + "  (version " + productVersion + ")");
    }

    static void CreateTables()
    {
        string[] tables =
        {
            "CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)",
            "CREATE TABLE `Directory` (`Directory` CHAR(72) NOT NULL, `Directory_Parent` CHAR(72), `DefaultDir` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Directory`)",
            "CREATE TABLE `Component` (`Component` CHAR(72) NOT NULL, `ComponentId` CHAR(38), `Directory_` CHAR(72) NOT NULL, `Attributes` SHORT NOT NULL, `Condition` CHAR(255), `KeyPath` CHAR(72) PRIMARY KEY `Component`)",
            "CREATE TABLE `File` (`File` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL LOCALIZABLE, `FileSize` LONG NOT NULL, `Version` CHAR(72), `Language` CHAR(20), `Attributes` SHORT, `Sequence` LONG NOT NULL PRIMARY KEY `File`)",
            "CREATE TABLE `Media` (`DiskId` SHORT NOT NULL, `LastSequence` LONG NOT NULL, `DiskPrompt` CHAR(64) LOCALIZABLE, `Cabinet` CHAR(255), `VolumeLabel` CHAR(32), `Source` CHAR(72) PRIMARY KEY `DiskId`)",
            "CREATE TABLE `Feature` (`Feature` CHAR(38) NOT NULL, `Feature_Parent` CHAR(38), `Title` CHAR(64) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`)",
            "CREATE TABLE `FeatureComponents` (`Feature_` CHAR(38) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Feature_`, `Component_`)",
            "CREATE TABLE `Registry` (`Registry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Registry`)",
            "CREATE TABLE `Shortcut` (`Shortcut` CHAR(72) NOT NULL, `Directory_` CHAR(72) NOT NULL, `Name` CHAR(128) NOT NULL LOCALIZABLE, `Component_` CHAR(72) NOT NULL, `Target` CHAR(72) NOT NULL, `Arguments` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `Hotkey` SHORT, `Icon_` CHAR(72), `IconIndex` SHORT, `ShowCmd` SHORT, `WkDir` CHAR(72) PRIMARY KEY `Shortcut`)",
            "CREATE TABLE `Icon` (`Name` CHAR(72) NOT NULL, `Data` OBJECT NOT NULL PRIMARY KEY `Name`)",
            "CREATE TABLE `Upgrade` (`UpgradeCode` CHAR(38) NOT NULL, `VersionMin` CHAR(20), `VersionMax` CHAR(20), `Language` CHAR(255), `Attributes` LONG NOT NULL, `Remove` CHAR(255), `ActionProperty` CHAR(72) NOT NULL PRIMARY KEY `UpgradeCode`, `VersionMin`, `VersionMax`, `Language`, `Attributes`)",
            "CREATE TABLE `CustomAction` (`Action` CHAR(72) NOT NULL, `Type` SHORT NOT NULL, `Source` CHAR(72), `Target` CHAR(255) PRIMARY KEY `Action`)",
            "CREATE TABLE `InstallExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)",
            "CREATE TABLE `InstallUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`)"
        };
        foreach (var sql in tables) Execute(sql);
    }

    static void WriteSummary()
    {
        IntPtr summary;
        Check(MsiGetSummaryInformation(db, null, 20, out summary), "MsiGetSummaryInformation");
        try
        {
            var none = new FILETIME();
            const uint VT_I2 = 2, VT_I4 = 3, VT_LPSTR = 30;
            Check(MsiSummaryInfoSetProperty(summary, 1, VT_I2, 1252, ref none, null), "codepage");
            Check(MsiSummaryInfoSetProperty(summary, 2, VT_LPSTR, 0, ref none, "Installation Database"), "title");
            Check(MsiSummaryInfoSetProperty(summary, 3, VT_LPSTR, 0, ref none, ProductName), "subject");
            Check(MsiSummaryInfoSetProperty(summary, 4, VT_LPSTR, 0, ref none, Manufacturer), "author");
            Check(MsiSummaryInfoSetProperty(summary, 7, VT_LPSTR, 0, ref none, "x64;1033"), "template");
            Check(MsiSummaryInfoSetProperty(summary, 9, VT_LPSTR, 0, ref none, NewGuid()), "package code");
            Check(MsiSummaryInfoSetProperty(summary, 14, VT_I4, 200, ref none, null), "schema");
            Check(MsiSummaryInfoSetProperty(summary, 15, VT_I4, 2, ref none, null), "word count");
            Check(MsiSummaryInfoSetProperty(summary, 18, VT_LPSTR, 0, ref none, "Outlook Notifier MsiBuilder"), "app name");
            Check(MsiSummaryInfoSetProperty(summary, 19, VT_I4, 2, ref none, null), "security");
            Check(MsiSummaryInfoPersist(summary), "MsiSummaryInfoPersist");
        }
        finally { MsiCloseHandle(summary); }
    }

    static void Insert(string table, string columns, params object[] values)
    {
        var cols = columns.Split(',');
        string sql = "INSERT INTO `" + table + "` (" + string.Join(", ", cols.Select(c => "`" + c + "`")) +
                     ") VALUES (" + string.Join(", ", cols.Select(c => "?")) + ")";
        Execute(sql, values);
    }

    static void Execute(string sql, params object[] values)
    {
        IntPtr view;
        Check(MsiDatabaseOpenView(db, sql, out view), sql);
        IntPtr record = IntPtr.Zero;
        try
        {
            if (values.Length > 0)
            {
                record = MsiCreateRecord((uint)values.Length);
                for (int i = 0; i < values.Length; i++)
                {
                    uint field = (uint)i + 1;
                    object v = values[i];
                    if (v == null) continue;
                    if (v is int) Check(MsiRecordSetInteger(record, field, (int)v), sql);
                    else if (v is FileData) Check(MsiRecordSetStream(record, field, ((FileData)v).Path), sql);
                    else Check(MsiRecordSetString(record, field, v.ToString()), sql);
                }
            }
            Check(MsiViewExecute(view, record), sql);
        }
        finally
        {
            MsiViewClose(view);
            MsiCloseHandle(view);
            if (record != IntPtr.Zero) MsiCloseHandle(record);
        }
    }

    static string MakeCabinet(string exe, string work)
    {
        string staging = Path.Combine(work, "cab");
        if (Directory.Exists(staging)) Directory.Delete(staging, true);
        Directory.CreateDirectory(staging);
        // Files inside the cabinet are named after their File table key.
        File.Copy(exe, Path.Combine(staging, "AppExe"));

        var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "makecab.exe"),
            "/D CompressionType=LZX /D CompressionMemory=21 /L \"" + staging + "\" AppExe app.cab")
        { WorkingDirectory = staging, UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
        using (var p = Process.Start(psi))
        {
            p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            if (p.ExitCode != 0) throw new Exception("makecab failed (" + p.ExitCode + ")");
        }
        return Path.Combine(staging, "app.cab");
    }

    static string GetFileVersion(string path)
    {
        var version = new StringBuilder(64);
        var lang = new StringBuilder(64);
        uint vLen = 64, lLen = 64;
        Check(MsiGetFileVersion(path, version, ref vLen, lang, ref lLen), "MsiGetFileVersion");
        return version.ToString();
    }

    static string NewGuid() { return Guid.NewGuid().ToString("B").ToUpperInvariant(); }

    static void Check(uint result, string what)
    {
        if (result != 0) throw new Exception(what + " -> code " + result);
    }
}
