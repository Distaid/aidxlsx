using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Exceptions;

/// <summary>
/// Some limit (generally set in <see cref="LoadOptions"/>) was exceeded.
/// </summary>
internal class LimitExceededException(string message) : XlsxException(message);
