# AGENTS.md

Guidance for coding agents working in this repository.

## Project Overview

WeihanLi.Npoi is a .NET library that provides NPOI-based helpers for importing and exporting Excel and CSV data. The public package targets `netstandard2.0` and exposes APIs for entity and `DataTable` import/export, attribute and Fluent API configuration, templates, multiple sheets, formatters, filters, freeze panes, and CSV handling.

Repository layout:

- `src/WeihanLi.Npoi`: library source.
- `test/WeihanLi.Npoi.Test`: xUnit v3 test project using Microsoft Testing Platform.
- `samples`: sample applications.
- `docs`: published docs and generated API/article content.
- `perf`: BenchmarkDotNet benchmarks.
- `build`: shared MSBuild props for versioning, package metadata, and signing.
- `build.cs`: C# based build script

## Environment

- Use the .NET SDK requested by `global.json` (`10.0.100`, with major roll-forward enabled).
- The root `Directory.Build.props` enables nullable reference types, implicit usings, and preview C# language features for most projects.
- Package versions are centrally managed in `Directory.Packages.props`; add or update package versions there instead of adding versions to individual `PackageReference` items.

## Build And Test Commands

Prefer the repo build script when validating changes:

```sh
dotnet build.cs
dotnet build.cs --target=build
dotnet build.cs --target=test
```

Useful direct commands:

```sh
dotnet restore WeihanLi.Npoi.slnx
dotnet build WeihanLi.Npoi.slnx
dotnet test test/WeihanLi.Npoi.Test/WeihanLi.Npoi.Test.csproj
```

For package-related changes, run the build script so package validation and shared build settings are exercised.

## Code Style Guidelines

- Follow `.editorconfig`.
- Use spaces; C# files use 4-space indentation, project/props XML files use 2-space indentation, and JSON uses 2-space indentation.
- C# files should include the configured file header:

  ```csharp
  // Copyright (c) Weihan Li. All rights reserved.
  // Licensed under the Apache license.
  ```

- Use file-scoped namespaces.
- Prefer `var` where the repo style already does.
- Keep `System` usings sorted according to the existing `.editorconfig` setting; do not force `System` first.
- Preserve nullable annotations and avoid suppressing warnings unless the reason is clear and local.
- Keep public API changes deliberate. This is a NuGet package with package validation enabled, so signature changes can affect consumers.
- Do not hand-edit generated resource designer output unless the resource file change requires regeneration.

## Testing Instructions

- Add or update tests in `test/WeihanLi.Npoi.Test` for behavior changes.
- Existing tests use xUnit v3 attributes such as `[Fact]` and `[Theory]`.
- Test data lives under `test/WeihanLi.Npoi.Test/TestData` and is copied to the test output. Add small, focused files only when necessary.
- Cover both Excel formats when the behavior is format-sensitive. The test project already has reusable Excel format theory data.
- For CSV changes, verify delimiters, header handling, quoting, null/empty values, and round-trip import/export behavior as applicable.
- For Excel import/export changes, verify both entity and `DataTable` paths when the touched code is shared.

## Security And Data Handling

- Treat spreadsheet and CSV input as untrusted. Avoid adding behavior that executes formulas, evaluates external links, expands paths from cell content, or writes files to caller-controlled paths without explicit API intent.
- Keep test data small and non-sensitive. Do not commit real customer spreadsheets, credentials, tokens, or private datasets.
- Avoid logging full cell contents or full input files in tests or samples when data may be user-provided.
- Do not weaken `NuGetAudit` settings in `Directory.Packages.props`.

## Documentation

- Update `README.md` and/or `docs/articles` when public behavior, usage patterns, or supported options change.
- The `docs/api` files are generated API reference content. Prefer regenerating them through the docs pipeline rather than manually editing generated API YAML.
- Keep examples concise and compatible with the package target and current public API.

## Pull Request Guidelines

- Keep changes focused and avoid unrelated formatting churn.
- Include tests for fixes and new behavior unless the change is documentation-only or build metadata-only.
- Mention any public API changes, package dependency changes, and compatibility considerations in the PR description.
- For commits, use short imperative messages, for example `Fix CSV empty column import` or `Add template export regression test`.

## Agent Workflow Notes

- Check `git status --short` before editing and avoid reverting user changes.
- Prefer `rg`/`rg --files` for repository search.
- Use `apply_patch` for manual edits.
- Before finishing, run the narrowest meaningful validation command for the change and report any commands that could not be run.
