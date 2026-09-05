using Serilog;
using Pfm;

try
{
    // Excel COM automation requires a single threaded apartment, and top level statements cannot carry [STAThread],
    // so the work runs on an STA thread that this one waits on.
    int exitCode = 0;
    var mainThread = new Thread(() => exitCode = Application.Run(args, Console.Out, Console.Error));
    mainThread.SetApartmentState(ApartmentState.STA);
    mainThread.Start();
    mainThread.Join();

    // Return so the finally block runs and flushes buffered log entries.
    return exitCode;
}
finally
{
    Log.CloseAndFlush();
}
