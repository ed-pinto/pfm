using System.Globalization;
using Microsoft.Extensions.Configuration;
using Serilog;

namespace Pfm;

public static class Application
{
    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        ArgumentNullException.ThrowIfNull(args);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);

        var commandLineParser = new CommandLineParser(args);

        InitializeLogging(commandLineParser.Configuration, commandLineParser.Arguments);

        if (commandLineParser.Error)
        {
            error.WriteLine(commandLineParser.ErrorMessage);
            Log.Logger.Error("PFM_INPUT_ERROR: " + commandLineParser.ErrorMessage);
            return 1;
        }
        else if (commandLineParser.Help)
        {
            output.WriteLine(commandLineParser.HelpMessage);
            return 0;
        }

        Log.Logger.Information("PFM_FILE_NAME: " + commandLineParser.Arguments.FilePath);

        return Dispatch(commandLineParser.Arguments, output, error);
    }

    /// <summary>
    /// Dispatches the parsed arguments to the operation that implements the requested command.
    /// </summary>
    /// <param name="arguments">The parsed command line arguments.</param>
    /// <param name="output">The writer for normal output.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>The exit code of the operation.</returns>
    private static int Dispatch(Arguments arguments, TextWriter output, TextWriter error)
    {
        return arguments.Command switch
        {
            Commands.Iterate => Operations.Iterate(arguments, output, error),
            Commands.ClosePeriod => NotImplemented(arguments.Command, error),
            Commands.MonteCarlo => NotImplemented(arguments.Command, error),
            Commands.BackTest => Operations.BackTest(arguments, output, error),
            _ => NotImplemented(arguments.Command, error)
        };
    }

    /// <summary>
    /// Reports a command that the CLI accepts but does not yet implement.
    /// </summary>
    /// <param name="command">The command that was requested.</param>
    /// <param name="error">The writer for error output.</param>
    /// <returns>One.</returns>
    private static int NotImplemented(Commands command, TextWriter error)
    {
        Log.Logger.Error("PFM_COMMAND_NOT_IMPLEMENTED: " + command);
        error.WriteLine("The command " + command + " has not been implemented.");
        return 1;
    }

    /// <summary>
    /// The path to the Serilog file sink's path setting within the configuration, expressed as a flattened
    /// configuration key.
    /// </summary>
    private const string FilePathConfigurationKey = "Serilog:WriteTo:0:Args:configure:0:Args:path";

    /// <summary>
    /// Initializes the logging system using the specified configuration.
    /// </summary>
    /// <param name="configuration">The configuration to use for logging.</param>
    /// <param name="args">
    /// The parsed command line arguments, or null if they are not available, ex. because parsing failed or help was
    /// requested. When a <see cref="Commands.BackTest"/> or <see cref="Commands.MonteCarlo"/> sweep is running, its
    /// job count and job index are folded into the log file name so that the jobs of one sweep do not overwrite each
    /// other's logs.
    /// </param>
    private static void InitializeLogging(IConfiguration configuration, Arguments? args)
    {
        // Serilog swallows sink write errors by default; surface them so misconfigurations are visible.
        Serilog.Debugging.SelfLog.Enable(msg => Console.Error.WriteLine("SERILOG_SELFLOG: " + msg));

        if (args is not null && (args.Command == Commands.BackTest || args.Command == Commands.MonteCarlo))
        {
            string? path = configuration[FilePathConfigurationKey];
            if (!string.IsNullOrEmpty(path))
            {
                string directory = Path.GetDirectoryName(path) ?? string.Empty;
                string fileName = Path.GetFileNameWithoutExtension(path);
                string extension = Path.GetExtension(path);
                string job = args.JobIndex.ToString(CultureInfo.InvariantCulture) + "_of_"
                    + args.JobCount.ToString(CultureInfo.InvariantCulture);
                string jobFileName = fileName + "." + job + extension;

                configuration = new ConfigurationBuilder()
                    .AddConfiguration(configuration)
                    .AddInMemoryCollection([new(FilePathConfigurationKey, Path.Combine(directory, jobFileName))])
                    .Build();
            }
        }

        LoggerConfiguration loggerConfig = new LoggerConfiguration()
            .ReadFrom.Configuration(configuration);
        Log.Logger = loggerConfig.CreateLogger();
    }

}
