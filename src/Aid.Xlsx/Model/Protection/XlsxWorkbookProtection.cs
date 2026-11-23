#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Utilities;
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;

namespace Aid.Xlsx.Model.Protection;

internal class XlsxWorkbookProtection(XlsxProtectionAlgorithm.Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements) : IXlsxWorkbookProtection
{
    public XlsxWorkbookProtection(Algorithm algorithm)
        : this(algorithm, XlsxWorkbookProtectionElements.Windows)
    {
    }

    public Algorithm Algorithm { get; internal set; } = algorithm;
    public XlsxWorkbookProtectionElements AllowedElements { get; set; } = allowedElements;
    public bool IsPasswordProtected => this.IsProtected && !string.IsNullOrEmpty(PasswordHash);
    public bool IsProtected { get; internal set; }

    internal string Base64EncodedSalt { get; set; }
    internal string PasswordHash { get; set; }
    internal uint SpinCount { get; set; } = 100000;

    public IXlsxWorkbookProtection AllowElement(XlsxWorkbookProtectionElements element, bool allowed = true)
    {
        if (allowed)
            AllowedElements |= element;
        else
            AllowedElements &= ~element;

        return this;
    }

    public IXlsxWorkbookProtection AllowEverything()
    {
        AllowedElements = XlsxWorkbookProtectionElements.Everything;
        return this;
    }

    public IXlsxWorkbookProtection AllowNone()
    {
        AllowedElements = XlsxWorkbookProtectionElements.None;
        return this;
    }

    public object Clone()
    {
        return new XlsxWorkbookProtection(this.Algorithm, this.AllowedElements)
        {
            IsProtected = this.IsProtected,
            PasswordHash = this.PasswordHash,
            SpinCount = this.SpinCount,
            Base64EncodedSalt = this.Base64EncodedSalt
        };
    }

    public IXlsxWorkbookProtection CopyFrom(IXlsxElementProtection<XlsxWorkbookProtectionElements> workbookProtection)
    {
        if (workbookProtection is XlsxWorkbookProtection xlWorkbookProtection)
        {
            this.IsProtected = xlWorkbookProtection.IsProtected;
            this.Algorithm = xlWorkbookProtection.Algorithm;
            this.PasswordHash = xlWorkbookProtection.PasswordHash;
            this.SpinCount = xlWorkbookProtection.SpinCount;
            this.Base64EncodedSalt = xlWorkbookProtection.Base64EncodedSalt;
            this.AllowedElements = xlWorkbookProtection.AllowedElements;
        }
        return this;
    }

    public IXlsxWorkbookProtection DisallowElement(XlsxWorkbookProtectionElements element)
    {
        return AllowElement(element, allowed: false);
    }

    public IXlsxWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protect(string.Empty, algorithm);
    }

    public IXlsxWorkbookProtection Protect(XlsxWorkbookProtectionElements allowedElements)
        => Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

    public IXlsxWorkbookProtection Protect(Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements)
        => Protect(string.Empty, algorithm, allowedElements);

    public IXlsxWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm, XlsxWorkbookProtectionElements allowedElements = XlsxWorkbookProtectionElements.Windows)
    {
        if (IsProtected)
        {
            throw new InvalidOperationException("The workbook structure is already protected");
        }
        else
        {
            IsProtected = true;

            password ??= "";

            this.Algorithm = algorithm;
            this.Base64EncodedSalt = CryptographicAlgorithms.GenerateNewSalt(this.Algorithm);
            this.PasswordHash = CryptographicAlgorithms.GetPasswordHash(this.Algorithm, password, this.Base64EncodedSalt, this.SpinCount);
        }

        this.AllowedElements = allowedElements;

        return this;
    }

    public IXlsxWorkbookProtection Unprotect()
    {
        return Unprotect(string.Empty);
    }

    public IXlsxWorkbookProtection Unprotect(string password)
    {
        if (IsProtected)
        {
            if (PasswordHash.Length > 0 && string.IsNullOrEmpty(password))
                throw new InvalidOperationException("The workbook structure is password protected");

            var hash = CryptographicAlgorithms.GetPasswordHash(this.Algorithm, password, this.Base64EncodedSalt, this.SpinCount);
            if (hash != PasswordHash)
                throw new ArgumentException("Invalid password");
            else
            {
                IsProtected = false;
                PasswordHash = string.Empty;
                this.Base64EncodedSalt = string.Empty;
            }
        }

        return this;
    }

    #region IXlsxProtectable interface

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.AllowElement(XlsxWorkbookProtectionElements element, bool allowed) => AllowElement(element, allowed);

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.AllowEverything() => AllowEverything();

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.AllowNone() => AllowNone();

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.DisallowElement(XlsxWorkbookProtectionElements element) => DisallowElement(element);

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.Protect(Algorithm algorithm) => Protect(algorithm);

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.Protect(string password, Algorithm algorithm) => Protect(password, algorithm);

    IXlsxWorkbookProtection IXlsxWorkbookProtection.Protect(XlsxWorkbookProtectionElements allowedElements) => Protect(allowedElements);

    IXlsxWorkbookProtection IXlsxWorkbookProtection.Protect(Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements) => Protect(algorithm, allowedElements);

    IXlsxWorkbookProtection IXlsxWorkbookProtection.Protect(string password, Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements) => Protect(password, algorithm, allowedElements);

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.Unprotect() => Unprotect();

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.Unprotect(string password) => Unprotect(password);

    IXlsxElementProtection<XlsxWorkbookProtectionElements> IXlsxElementProtection<XlsxWorkbookProtectionElements>.CopyFrom(IXlsxElementProtection<XlsxWorkbookProtectionElements> protectable) => CopyFrom(protectable);

    #endregion IXlsxProtectable interface
}
