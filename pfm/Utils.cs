using System.Diagnostics.CodeAnalysis;

namespace Pfm;

public static class Utils
{
    /// <summary>
    /// Validates whether the specified path is structurally valid and points to an Excel file.
    /// </summary>
    /// <param name="path">The path to validate.</param>
    /// <returns>True if the path is valid and points to an Excel file; otherwise, false.</returns>
    public static bool IsValidExcelPath([NotNullWhen(true)] string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return false;

        // Reject invalid characters (null chars, control chars, etc.)
        if (path.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            return false;

        try
        {
            // GetFullPath will throw for structurally invalid paths
            // (e.g. malformed UNC paths, invalid drive syntax, paths that are too long, etc.)
            string fullPath = Path.GetFullPath(path);

            // Optional: also validate the file name portion, since
            // GetInvalidPathChars() doesn't include everything GetInvalidFileNameChars() catches
            string fileName = Path.GetFileName(fullPath);
            if (!string.IsNullOrEmpty(fileName) &&
                fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return false;
            }

            // Verify the file has an xlsx extension
            if (!string.IsNullOrEmpty(fileName) && !fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Ensure the file exists
            if (!File.Exists(fullPath))
                return false;

            return true;
        }
        catch (ArgumentException) { return false; }
        catch (PathTooLongException) { return false; }
        catch (NotSupportedException) { return false; } // e.g. malformed drive/colon usage
        catch (System.Security.SecurityException) { return false; }
    }
}
