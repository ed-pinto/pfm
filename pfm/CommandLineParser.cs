using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Extensions.Configuration;

namespace Pfm;

/// <summary>
/// Represents the command line interface for pfm, handling argument parsing and preparation.
/// </summary>
public class CommandLineParser
{
    private const string Pfm = "pfm";
    private const string IterateCommand = "iterate";
    private const string ClosePeriodCommand = "close-period";
    private const string MonteCarloCommand = "monte-carlo";
    private const string BackTestCommand = "back-test";

    private bool _error;
    private string? _errorMessage;
    private bool _help;
    private string? _helpMessage;
    private Commands _command;
    private string? _filePath;
    private bool _timing;

    private Arguments? _arguments;

    /// <summary>
    /// Initializes a new instance of the <see cref="CommandLineParser"/> class, parsing the provided command line arguments
    /// and preparing the environment and settings as needed.
    /// </summary>
    /// <param name="args">The command line arguments to parse.</param>
    /// <remarks>
    /// After construction, the <see cref="CommandLineParser"/> instance will have parsed the arguments.
    /// The Arguments property will be initialized if there were no errors (Error is false) and the help option was not
    /// specified (Help is false).
    /// If there were errors during parsing, the <see cref="Error"/> property will be true and the
    /// <see cref="ErrorMessage"/> property will be set.
    /// If the help option was specified, the <see cref="Help"/> property will be set to true and the
    /// <see cref="HelpMessage"/> property will be set with the help information.
    /// </remarks>
    public CommandLineParser(string[] args)
    {
        ParseArguments(args);

        Configuration = BuildConfiguration();

        if (!Error && !Help)
        {
            PrepareArguments();
        }
    }

    /// <summary>
    /// Gets the configuration object built from the environment-specific configuration files.
    /// </summary>
    public IConfiguration Configuration { get; init; }

    /// <summary>
    /// Indicates whether the help option was specified in the command line arguments.  Settings will be null when this
    /// property is true.
    /// </summary>
    public bool Help => _help;

    /// <summary>
    /// Gets the help message generated for the command line arguments.  This will be null if Help is false.
    /// </summary>
    public string? HelpMessage => _helpMessage;

    /// <summary>
    /// Indicates whether an error occurred during command line parsing.  Settings will be null when this property is true.
    /// </summary>
    [MemberNotNullWhen(true, nameof(ErrorMessage))]
    public bool Error => _error;

    /// <summary>
    /// Gets the error message associated with a command line parsing error.  This will be null if Error is false.
    /// </summary>
    public string? ErrorMessage => _errorMessage;

    /// <summary>
    /// Gets the parsed command line arguments.  This will be null if there were errors or if the help option was specified.
    /// </summary>
    public Arguments Arguments => _arguments ?? throw new InvalidOperationException("Arguments have not been initialized. "
        + "Either the Error or Help property is true.");

    private void ParseArguments(string[] args)
    {
        var rootCommand = new RootCommand();

        var iterateCommand = new Command(IterateCommand, "Iterates calculations in the PortfolioManager workbook which "
            + "require iterative updates.");
        Option<string> filePathOption = new("--file-path")
        {
            Description = "Specifies the path to the PortfolioManager workbook.",
            Recursive = true,
            Required = true,
            Aliases = { "-f" }
        };

        Option<bool> timingOption = new("--timing")
        {
            Description = "Reports how long each phase of the command takes, ex. opening and saving the workbook.",
            Recursive = true,
            Required = false,
            Aliases = { "-t" }
        };

        rootCommand.Options.Add(filePathOption);
        rootCommand.Options.Add(timingOption);
        rootCommand.Subcommands.Add(iterateCommand);

        rootCommand.Options.Remove(rootCommand.Options.OfType<VersionOption>().Single());
        HelpOption helpOption = rootCommand.Options.OfType<HelpOption>().Single();

        ParseResult parseResult = rootCommand.Parse(args);

        if (parseResult.Errors.Count > 0)
        {
            string errorMessage = string.Join(Environment.NewLine, parseResult.Errors);
            SetError(errorMessage);
        }
        else if (parseResult.GetResult(helpOption) != null)
        {
            _help = true;
            var helpWriter = new StringWriter();
            rootCommand.Parse("-h").Invoke(new InvocationConfiguration { Output = helpWriter });
            _helpMessage = helpWriter.ToString();
        }
        else
        {
            switch (parseResult.CommandResult.Command.Name)
            {
                case IterateCommand:
                    _command = Commands.Iterate;
                    break;
                default:
                    SetError("Unknown command: " + parseResult.CommandResult.Command.Name);
                    break;
            }

            _filePath = parseResult.GetValue(filePathOption);
            _timing = parseResult.GetValue(timingOption);
        }
    }

    private void PrepareArguments()
    {
        if (!Utils.IsValidExcelPath(_filePath))
        {
            SetError("Invalid file path: " + _filePath + ". The file must be a valid Excel (.xlsx) file.");
            return;
        }

        _arguments = new Arguments
        {
            Command = _command,
            FilePath = _filePath,
            Timing = _timing
        };
    }

    /// <summary>
    /// Sets the error state and the associated error message for the command line parser.
    /// </summary>
    /// <param name="errorMessage">The error message to set.</param>
    private void SetError(string errorMessage)
    {
        _error = true;
        _errorMessage = errorMessage;
    }

    /// <summary>
    /// Builds the application configuration.
    /// </summary>
    /// <returns>The built <see cref="IConfiguration"/> instance.</returns>
    private static IConfiguration BuildConfiguration()
    {
        // The settings files are optional.  No exception should be thrown if they are missing.  Missing configuration
        // that is required should be validated where it is read (ex. when initializing a Settings object)..
        IConfigurationBuilder builder = new ConfigurationBuilder()
            .AddJsonFile("settings.json", optional: true);

        return builder.Build();
    }

}
