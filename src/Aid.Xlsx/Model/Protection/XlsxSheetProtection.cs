#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Utilities;
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;

namespace Aid.Xlsx.Model.Protection;

internal class XlsxSheetProtection(XlsxProtectionAlgorithm.Algorithm algorithm) : IXlsxSheetProtection
{
    public Algorithm Algorithm { get; internal set; } = algorithm;
    public XlsxSheetProtectionElements AllowedElements { get; set; } = XlsxSheetProtectionElements.SelectEverything;

    public Boolean IsPasswordProtected => this.IsProtected && !String.IsNullOrEmpty(PasswordHash);
    public Boolean IsProtected { get; internal set; }

    internal String Base64EncodedSalt { get; set; }
    internal String PasswordHash { get; set; }
    internal UInt32 SpinCount { get; set; } = 100000;

    public IXlsxSheetProtection AllowElement(XlsxSheetProtectionElements element, Boolean allowed = true)
    {
        if (!allowed)
            return DisallowElement(element);

        AllowedElements |= element;
        return this;
    }

    public IXlsxSheetProtection AllowEverything()
    {
        return AllowElement(XlsxSheetProtectionElements.Everything);
    }

    public IXlsxSheetProtection AllowNone()
    {
        AllowedElements = XlsxSheetProtectionElements.None;
        return this;
    }

    public object Clone()
    {
        return new XlsxSheetProtection(this.Algorithm)
        {
            IsProtected = this.IsProtected,
            PasswordHash = this.PasswordHash,
            SpinCount = this.SpinCount,
            Base64EncodedSalt = this.Base64EncodedSalt,
            AllowedElements = this.AllowedElements
        };
    }

    public XlsxSheetProtection CopyFrom(IXlsxElementProtection<XlsxSheetProtectionElements> sheetProtection)
    {
        if (sheetProtection is XlsxSheetProtection xlSheetProtection)
        {
            this.IsProtected = xlSheetProtection.IsProtected;
            this.Algorithm = xlSheetProtection.Algorithm;
            this.PasswordHash = xlSheetProtection.PasswordHash;
            this.SpinCount = xlSheetProtection.SpinCount;
            this.Base64EncodedSalt = xlSheetProtection.Base64EncodedSalt;
            this.AllowedElements = xlSheetProtection.AllowedElements;
        }
        return this;
    }

    public IXlsxSheetProtection DisallowElement(XlsxSheetProtectionElements element)
    {
        AllowedElements &= ~element;
        return this;
    }

    public IXlsxSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protect(String.Empty, algorithm);
    }

    public IXlsxSheetProtection Protect(XlsxSheetProtectionElements allowedElements)
        => Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

    public IXlsxSheetProtection Protect(Algorithm algorithm, XlsxSheetProtectionElements allowedElements)
        => Protect(string.Empty, algorithm, allowedElements);

    public IXlsxSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XlsxSheetProtectionElements allowedElements = XlsxSheetProtectionElements.SelectEverything)
    {
        if (IsProtected)
        {
            throw new InvalidOperationException("The worksheet is already protected");
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

    public IXlsxSheetProtection Unprotect()
    {
        return Unprotect(String.Empty);
    }

    public IXlsxSheetProtection Unprotect(String password)
    {
        if (IsProtected)
        {
            if (PasswordHash.Length > 0 && string.IsNullOrEmpty(password))
                throw new InvalidOperationException("The worksheet is password protected");

            var hash = CryptographicAlgorithms.GetPasswordHash(this.Algorithm, password, this.Base64EncodedSalt, this.SpinCount);
            if (hash != PasswordHash)
                throw new ArgumentException("Invalid password");
            else
            {
                IsProtected = false;
                PasswordHash = String.Empty;
                this.Base64EncodedSalt = String.Empty;
            }
        }

        return this;
    }

    #region IXlsxProtectable interface

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.AllowElement(XlsxSheetProtectionElements element, Boolean allowed) => AllowElement(element, allowed);

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.AllowEverything() => AllowEverything();

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.AllowNone() => AllowNone();

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.CopyFrom(IXlsxElementProtection<XlsxSheetProtectionElements> protectable) => CopyFrom(protectable);

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.DisallowElement(XlsxSheetProtectionElements element) => DisallowElement(element);

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.Protect(Algorithm algorithm) => Protect(algorithm);

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.Protect(String password, Algorithm algorithm) => Protect(password, algorithm);

    IXlsxSheetProtection IXlsxSheetProtection.Protect(XlsxSheetProtectionElements allowedElements) => Protect(allowedElements);

    IXlsxSheetProtection IXlsxSheetProtection.Protect(Algorithm algorithm, XlsxSheetProtectionElements allowedElements) => Protect(algorithm, allowedElements);

    IXlsxSheetProtection IXlsxSheetProtection.Protect(String password, Algorithm algorithm, XlsxSheetProtectionElements allowedElements) => Protect(password, algorithm, allowedElements);

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.Unprotect() => Unprotect();

    IXlsxElementProtection<XlsxSheetProtectionElements> IXlsxElementProtection<XlsxSheetProtectionElements>.Unprotect(String password) => Unprotect(password);

    #endregion IXlsxProtectable interface
}
