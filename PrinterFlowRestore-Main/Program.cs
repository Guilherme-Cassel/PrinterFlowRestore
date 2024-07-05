using System.Security.Principal;
using System.ServiceProcess;

namespace ConsoleTesting2;

internal class Program
{
    public static ServiceController PrinterSpooler = new ServiceController("Spooler de Impressão");

    static void Main()
    {
        if (!IsRunningAsAdministrator())
        {
            TerminateConsole("Please run the software as Administrator", ConsoleColor.Blue);
        }

        try
        {
            StopService();
            CleanSpoolQueueFolder();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error While Cleaning Printer Queues: {ex.Message}");
        }
        finally
        {
            StartService();
            TerminateConsole($"Printer Queues Cleaned Successfully", ConsoleColor.Green);
        }
    }

    public static void TerminateConsole(string message, ConsoleColor color = ConsoleColor.White)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(message);
        Console.ForegroundColor = ConsoleColor.White;
        Console.ReadLine();
        Environment.Exit(0);
    }

    public static void StopService()
    {
        PrinterSpooler.Stop();
        PrinterSpooler.WaitForStatus(ServiceControllerStatus.Stopped);
    }

    public static void StartService()
    {
        PrinterSpooler.Start();
        PrinterSpooler.WaitForStatus(ServiceControllerStatus.Running);
    }

    public static void CleanSpoolQueueFolder()
    {
        string path = "C:\\Windows\\System32\\spool\\PRINTERS\\";

        foreach (string entry in Directory.GetFileSystemEntries(path))
        {
            if (File.Exists(entry))
            {
                File.Delete(entry);
            }
            else if (Directory.Exists(entry))
            {
                Directory.Delete(entry, true);
            }
        }
    }

    public static bool IsRunningAsAdministrator()
    {
        using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
        {
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}