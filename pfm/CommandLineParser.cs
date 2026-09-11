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
    private const string CoalesceCommand = "coalesce";

    /// <summary>
    /// The input path a coalesce uses when the caller does not give one: the current directory, which is where a
    /// sweep leaves its run directories.
    /// </summary>
    private const string DefaultInputPath = "./";

    /// <summary>
    /// The output path a coalesce uses when the caller does not give one.
    /// </summary>
    private const string DefaultOutputPath = "./output";

    private bool _error;
    private string? _errorMessage;
    private bool _help;
    private string? _helpMessage;
    private Commands _command;
    private string? _filePath;
    private bool _timing;
    private int _jobCount = 1;
    private int _jobIndex;
    private int? _simulationCount;
    private string _inputPath = DefaultInputPath;
    private string _outputPath = DefaultOutputPath;

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

    /// <summary>
    /// Gets the parsed command line arguments, or null when there are none because parsing failed or help was
    /// requested.  This is what a caller that runs before those outcomes are handled reads, ex. one initializing
    /// logging.
    /// </summary>
    public Arguments? ArgumentsOrDefault => _arguments;

    private void ParseArguments(string[] args)
    {
        var rootCommand = new RootCommand();

        var iterateCommand = new Command(IterateCommand, "Iterates calculations in the PortfolioManager workbook which "
            + "require iterative updates.");

        // The workbook option belongs to the commands that drive a workbook rather than to the root, because coalesce
        // has no workbook to be pointed at: it reads what a finished sweep left on disk.  It is still declared once
        // and added to each of those commands, so they describe it identically and read it from one place.
        Option<string> filePathOption = new("--file-path")
        {
            Description = "Specifies the path to the PortfolioManager workbook.",
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

        var backTestCommand = new Command(BackTestCommand, "Projects the PortfolioManager workbook from every "
            + "historical period long enough to project over, and records the outcome of each.");

        var monteCarloCommand = new Command(MonteCarloCommand, "Projects the PortfolioManager workbook down every "
            + "Monte Carlo path 22_MCSeeds holds seeds for, and records the outcome of each.");

        // The job options belong to the sweeping commands rather than to the root: iterating or closing a period is
        // one indivisible piece of work, so there is nothing for a slice of it to mean.
        Option<int> jobCountOption = new("--job-count")
        {
            Description = "The number of jobs the sweep is being split across.  Defaults to one, which runs it whole.",
            Required = false,
            DefaultValueFactory = _ => 1
        };

        Option<int> jobIndexOption = new("--job-index")
        {
            Description = "The zero based index of the slice of the sweep this process runs.  Defaults to zero.",
            Required = false,
            DefaultValueFactory = _ => 0
        };

        // Absent rather than defaulted, because there is no number of simulations that means every one of them: the
        // count a sweep offers is the workbook's to state, and this only ever shortens it.
        Option<int?> simulationCountOption = new("--simulation-count")
        {
            Description = "The greatest number of simulations the whole sweep runs, ex. 100 to run the first hundred "
                + "of a thousand iteration campaign.  Defaults to every simulation the workbook offers.",
            Required = false
        };

        backTestCommand.Options.Add(jobCountOption);
        backTestCommand.Options.Add(jobIndexOption);
        backTestCommand.Options.Add(simulationCountOption);
        monteCarloCommand.Options.Add(jobCountOption);
        monteCarloCommand.Options.Add(jobIndexOption);
        monteCarloCommand.Options.Add(simulationCountOption);

        var coalesceCommand = new Command(CoalesceCommand, "Gathers the run directories of one simulation sweep into a "
            + "single Excel workbook.");

        Option<string> inputPathOption = new("--input-path")
        {
            Description = "The directory holding the run directories of the sweep.  Defaults to the current directory.",
            Required = false,
            Aliases = { "-i" },
            DefaultValueFactory = _ => DefaultInputPath
        };

        Option<string> outputPathOption = new("--output-path")
        {
            Description = "The directory the coalesced workbook is written to.  Defaults to " + DefaultOutputPath
                + ".",
            Required = false,
            Aliases = { "-o" },
            DefaultValueFactory = _ => DefaultOutputPath
        };

        coalesceCommand.Options.Add(inputPathOption);
        coalesceCommand.Options.Add(outputPathOption);

        iterateCommand.Options.Add(filePathOption);
        backTestCommand.Options.Add(filePathOption);
        monteCarloCommand.Options.Add(filePathOption);

        rootCommand.Options.Add(timingOption);
        rootCommand.Subcommands.Add(iterateCommand);
        rootCommand.Subcommands.Add(backTestCommand);
        rootCommand.Subcommands.Add(monteCarloCommand);
        rootCommand.Subcommands.Add(coalesceCommand);

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

            // Invoking the parse result rather than a fresh one for the root command is what makes "back-test -h"
            // describe the back test command's own options, ex. --job-count, rather than only the root's.
            parseResult.Invoke(new InvocationConfiguration { Output = helpWriter });
            _helpMessage = helpWriter.ToString();
        }
        else
        {
            switch (parseResult.CommandResult.Command.Name)
            {
                case IterateCommand:
                    _command = Commands.Iterate;
                    break;
                case BackTestCommand:
                case MonteCarloCommand:
                    _command = parseResult.CommandResult.Command.Name == BackTestCommand
                        ? Commands.BackTest
                        : Commands.MonteCarlo;

                    // The job options belong to the sweeping commands alone, so they are only in scope to be read
                    // here.  Asking for one that the parsed command does not declare yields the type's default rather
                    // than the option's, which for a job count would be zero.
                    _jobCount = parseResult.GetValue(jobCountOption);
                    _jobIndex = parseResult.GetValue(jobIndexOption);
                    _simulationCount = parseResult.GetValue(simulationCountOption);
                    break;
                case CoalesceCommand:
                    _command = Commands.Coalesce;

                    // The options declare their own defaults, so these are never null for the coalesce command.
                    _inputPath = parseResult.GetValue(inputPathOption) ?? DefaultInputPath;
                    _outputPath = parseResult.GetValue(outputPathOption) ?? DefaultOutputPath;
                    break;
                default:
                    SetError("Unknown command: " + parseResult.CommandResult.Command.Name);
                    break;
            }

            // The workbook option is declared by the commands that drive one, so this yields null for a command that
            // does not, which is what PrepareArguments then declines to validate as a path.
            _filePath = parseResult.GetValue(filePathOption);
            _timing = parseResult.GetValue(timingOption);
        }
    }

    private void PrepareArguments()
    {
        // Coalesce reads what a finished sweep left on disk rather than driving a workbook, so it is the one command
        // with no workbook to validate.  Its own paths are directories, one of which does not exist yet, so they are
        // checked where they are used rather than for existence here.
        if (_command != Commands.Coalesce && !Utils.IsValidExcelPath(_filePath))
        {
            SetError("Invalid file path: " + _filePath + ". The file must be a valid Excel (.xlsx) file.");
            return;
        }

        if (_command == Commands.Coalesce && !Utils.IsValidDirectoryPath(_inputPath))
        {
            SetError("Invalid input path: " + _inputPath + ". It must be a directory that exists.");
            return;
        }

        if (_command == Commands.Coalesce && !Utils.IsValidNewDirectoryPath(_outputPath))
        {
            SetError("Invalid output path: " + _outputPath + ". It must be a path a directory can be created at.");
            return;
        }

        if (_jobCount < 1)
        {
            SetError("Invalid job count: " + Format(_jobCount) + ". A sweep must be split across at least one job.");
            return;
        }

        if (_jobIndex < 0 || _jobIndex >= _jobCount)
        {
            SetError("Invalid job index: " + Format(_jobIndex) + ". The index must be at least zero and less than the "
                + "job count of " + Format(_jobCount) + ".");
            return;
        }

        if (_simulationCount is int simulationCount && simulationCount < 1)
        {
            SetError("Invalid simulation count: " + Format(simulationCount) + ". A sweep must run at least one "
                + "simulation.  Leave the option out to run every simulation the workbook offers.");
            return;
        }

        _arguments = new Arguments
        {
            Command = _command,
            FilePath = _filePath,
            Timing = _timing,
            JobCount = _jobCount,
            JobIndex = _jobIndex,
            SimulationCount = _simulationCount,
            InputPath = _inputPath,
            OutputPath = _outputPath
        };
    }

    /// <summary>
    /// Formats a count for an error message.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <returns>The formatted value.</returns>
    private static string Format(int value)
    {
        return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
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
