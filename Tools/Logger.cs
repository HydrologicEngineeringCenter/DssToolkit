using Microsoft.Extensions.Logging.Console;
using Microsoft.Extensions.Logging;
public class Logger
{
  private static Logger onlyInstance;
  private readonly ILogger logger1;

  private Logger()
  {
    var loggerFactory = LoggerFactory.Create(builder =>
    {
    builder.AddConsole(); // add other logging providers here
    });

    logger1 = loggerFactory.CreateLogger<Logger>();
  }

  public static Logger Instance
  {
    get
    {
      if (onlyInstance == null)
      {
        onlyInstance = new Logger();
      }
      return onlyInstance;
    }
  }

  public static void LogMessage(string message)
  {
    Instance.logger1.LogInformation(message);
  }

  public static void LogError(string errorMessage)
  {
    Instance.logger1.LogError(errorMessage);
  }
}
