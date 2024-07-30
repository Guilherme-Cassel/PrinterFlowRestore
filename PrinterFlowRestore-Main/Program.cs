using System.Security.Principal;
using System.ServiceProcess;

namespace PrinterFlowRestore;

public class Program
{
    public static ServiceController PrinterSpooler = new ServiceController("Spooler de Impressão");
    public static string SpoolerJobsPath = "C:\\Windows\\System32\\spool\\PRINTERS\\";

    static void Main()
    {
        if (!IsRunningAsAdministrator())
        {
            Log("Please run the software as Administrator", ConsoleColor.Blue);
            TerminateConsole();
        }

        try
        {
            StopService();
            CleanSpoolQueueFolder();
            Log($"Printer Queues Cleaned Successfully", ConsoleColor.Green);
        }
        catch (Exception ex)
        {
            Log($"Error While Cleaning Printer Queues: {ex.Message}", ConsoleColor.Red);
        }
        finally
        {
            RestartService();
            TerminateConsole();
        }
    }

    public static void TerminateConsole()
    {
        Console.ReadLine();
        Environment.Exit(0);
    }

    public static void StopService()
    {
        try
        {
            PrinterSpooler.Stop();
            PrinterSpooler.WaitForStatus(ServiceControllerStatus.Stopped);
        }
        catch (InvalidOperationException)
        {
            //the service is already stopped.
        }
        finally
        {
            Log("Spooler Service Stopped", ConsoleColor.Blue);
        }
    }

    public static void RestartService()
    {
        try
        {
            PrinterSpooler.Start();
            PrinterSpooler.WaitForStatus(ServiceControllerStatus.Running);
        }
        catch (InvalidOperationException)
        {
            //the service is already running.
        }
        finally
        {
            Log("Spooler Service Restarted", ConsoleColor.Blue);
        }
    }

    public static void CleanSpoolQueueFolder()
    {
        foreach (string entry in Directory.GetFileSystemEntries(SpoolerJobsPath))
        {
            Directory.Delete(entry, true);
        }
    }

    public static bool IsRunningAsAdministrator()
    {
        using WindowsIdentity identity = WindowsIdentity.GetCurrent();
        WindowsPrincipal principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    public static void Log(string msg, ConsoleColor color)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(msg);
        Console.ForegroundColor = ConsoleColor.White;
    }
}