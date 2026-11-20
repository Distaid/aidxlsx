# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Restructured the solution for the port: the projects and namespaces are
  `Aid.Xlsx.Opc` (OPC/XML layer) and `Aid.Xlsx.CodeGen` (XSD code generator); the core
  library now uses a flat layout вЂ” domain folders sit directly under `src/Aid.Xlsx/`,
  part readers and writers live in `src/Aid.Xlsx/Internal/IO/`.

### Added

- Initial repository files: `.gitignore`, the solution file `Aid.Xlsx.slnx` (currently
  without projects) and the project documentation (`README.md`, this changelog).

### Planned

- `0.1.0-preview1` — the first public preview on NuGet, planned for after the upstream
  import, the custom OPC/XML layer and the writer migration are complete. The version is
  planned, not released: no release date is set and no package exists yet.

[Unreleased]: https://github.com/Distaid/aidxlsx/commits/main