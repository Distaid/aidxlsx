namespace Aid.Xlsx.Calculation;

/// <summary>
/// The exception that is thrown when the strings to be parsed to an expression is invalid.
/// </summary>
/// <remarks>
/// Initializes a new instance of the ExpressionParseException class with a
/// specified error message.
/// </remarks>
/// <param name="message">The message that describes the error.</param>
public class ExpressionParseException(string message) : Exception(message)
{
}
