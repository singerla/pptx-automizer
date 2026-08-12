using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

// ROADMAP Phase 5, Tier 2 — OOXML schema validation.
//
// Validates .pptx files with the Open XML SDK validator — the genuine oracle
// for the "PowerPoint repair prompt" bug class (schema child order etc.) that
// XML assertions, package invariants and pixel diffs all miss.
//
// Usage:  validate-pptx [--allowlist allowlist.json] <file-or-dir>...
//
// Directories are expanded to their *.pptx files (non-recursive). Errors
// matching an allowlist entry are counted but do not fail the run — the
// allowlist is seeded from the pre-existing warnings in the test templates
// (real-world files carry some), so only *new* error classes fail. Exit code
// 0 = no new errors.

var allowlistPath = "";
var verbose = false;
var targets = new List<string>();
for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--allowlist":
            allowlistPath = args[++i];
            break;
        case "--verbose":
            verbose = true;
            break;
        default:
            targets.Add(args[i]);
            break;
    }
}

if (targets.Count == 0)
{
    Console.Error.WriteLine("usage: validate-pptx [--allowlist allowlist.json] [--verbose] <file-or-dir>...");
    return 2;
}

var allowlist = allowlistPath == ""
    ? new List<AllowlistEntry>()
    : JsonSerializer.Deserialize<List<AllowlistEntry>>(
          File.ReadAllText(allowlistPath),
          new JsonSerializerOptions { PropertyNameCaseInsensitive = true, ReadCommentHandling = JsonCommentHandling.Skip })
      ?? new List<AllowlistEntry>();

var files = targets
    .SelectMany(target => Directory.Exists(target)
        ? Directory.EnumerateFiles(target, "*.pptx").OrderBy(file => file, StringComparer.Ordinal)
        : (IEnumerable<string>)new[] { target })
    .ToList();

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var newErrors = 0;
var allowlisted = 0;

foreach (var path in files)
{
    try
    {
        using var document = PresentationDocument.Open(path, false);
        foreach (var error in validator.Validate(document))
        {
            var partUri = error.Part?.Uri.ToString() ?? "";
            if (allowlist.Any(entry => entry.Matches(error.Description, partUri)))
            {
                allowlisted++;
                if (verbose)
                {
                    Console.WriteLine($"{Path.GetFileName(path)} :: allowlisted :: [{error.ErrorType}] {error.Description} @ {partUri} {error.Path?.XPath}");
                }
                continue;
            }
            newErrors++;
            Console.WriteLine($"{Path.GetFileName(path)} :: [{error.ErrorType}] {error.Description} @ {partUri} {error.Path?.XPath}");
        }
    }
    catch (Exception exception)
    {
        // A file the SDK cannot even open is the worst validation result —
        // but the allowlist applies here too (e.g. the known dangling-rels
        // classes make PresentationDocument.Open throw).
        if (allowlist.Any(entry => entry.Matches(exception.Message, "")))
        {
            allowlisted++;
            if (verbose)
            {
                Console.WriteLine($"{Path.GetFileName(path)} :: allowlisted :: [OpenFailed] {exception.Message}");
            }
            continue;
        }
        newErrors++;
        Console.WriteLine($"{Path.GetFileName(path)} :: [OpenFailed] {exception.Message}");
    }
}

Console.WriteLine($"validate-pptx: {files.Count} files checked, {newErrors} new errors, {allowlisted} allowlisted");
return newErrors == 0 ? 0 : 1;

record AllowlistEntry(string Description, string? Part)
{
    public bool Matches(string description, string partUri) =>
        description.Contains(Description, StringComparison.Ordinal) &&
        (Part is null || partUri.Contains(Part, StringComparison.Ordinal));
}
