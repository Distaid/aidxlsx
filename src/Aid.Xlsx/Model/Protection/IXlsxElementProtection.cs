#nullable disable

// Keep this file CodeMaid organised and cleaned
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;

namespace Aid.Xlsx.Model.Protection;

public interface IXlsxElementProtection<T> : IXlsxElementProtection
    where T : struct
{
    /// <summary>Gets or sets the elements that are allowed to be edited by the user, i.e. those that are not protected.</summary>
    /// <value>The allowed elements.</value>
    T AllowedElements { get; set; }

    /// <summary>
    /// Adds the specified element to the list of allowed elements.
    /// Beware that if you pass through "None", this will have no effect.
    /// </summary>
    /// <param name="element">The element to add</param>
    /// <param name="allowed">Set to <c>true</c> to allow the element or <c>false</c> to disallow the element</param>
    /// <returns>The current protection instance</returns>
    IXlsxElementProtection<T> AllowElement(T element, Boolean allowed = true);

    /// <summary>Allows all elements to be edited.</summary>
    IXlsxElementProtection<T> AllowEverything();

    /// <summary>Allows no elements to be edited. Protects all elements.</summary>
    IXlsxElementProtection<T> AllowNone();

    /// <summary>Copies all the protection settings from a different instance.</summary>
    /// <param name="protectable">The protectable.</param>
    IXlsxElementProtection<T> CopyFrom(IXlsxElementProtection<T> protectable);

    /// <summary>
    /// Removes the element to the list of allowed elements.
    /// Beware that if you pass through "None", this will have no effect.
    /// </summary>
    /// <param name="element">The element to remove</param>
    /// <returns>The current protection instance</returns>
    IXlsxElementProtection<T> DisallowElement(T element);

    /// <summary>Protects this instance without a password.</summary>
    /// <param name="algorithm">The algorithm.</param>
    IXlsxElementProtection<T> Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

    /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
    /// <param name="password">The password.</param>
    /// <param name="algorithm">The algorithm.</param>
    IXlsxElementProtection<T> Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm);

    /// <summary>Unprotects this instance without a password.</summary>
    IXlsxElementProtection<T> Unprotect();

    /// <summary>Unprotects this instance using the specified password.</summary>
    /// <param name="password">The password.</param>
    IXlsxElementProtection<T> Unprotect(String password);
}

public interface IXlsxElementProtection : ICloneable
{
    /// <summary>Gets the algorithm used to hash the password.</summary>
    /// <value>The algorithm.</value>
    Algorithm Algorithm { get; }

    /// <summary>Gets a value indicating whether this instance is protected with a password.</summary>
    /// <value>
    ///   <c>true</c> if this instance is password protected; otherwise, <c>false</c>.
    /// </value>
    Boolean IsPasswordProtected { get; }

    /// <summary>Gets a value indicating whether this instance is protected, either with or without a password.</summary>
    /// <value>
    ///   <c>true</c> if this instance is protected; otherwise, <c>false</c>.
    /// </value>
    Boolean IsProtected { get; }
}
