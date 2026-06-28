using System.Globalization;
using System.Text.RegularExpressions;
using ExcelPerfDiag.Core.Model;

namespace ExcelPerfDiag.Core.Analysis;

/// <summary>
/// Pure formula analysis (no Excel): classification, R1C1 normalisation, the
/// (R1C1, cells-touched bucket) grouping key, cells-touched estimate and
/// full-column detection. Faithful port of the verified Python <c>patterns.py</c>.
/// </summary>
public static partial class FormulaAnalysis
{
    // Function-name sets (the cost-driver taxonomy).
    private static readonly HashSet<string> Volatile = new(StringComparer.Ordinal)
        { "OFFSET", "INDIRECT", "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "CELL", "INFO" };
    private static readonly HashSet<string> SingleThreaded = new(StringComparer.Ordinal)
        { "INDIRECT", "GETPIVOTDATA", "CELL", "INFO" };
    private static readonly HashSet<string> Lookup = new(StringComparer.Ordinal)
        { "VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "INDEX", "XLOOKUP", "XMATCH" };
    private static readonly HashSet<string> ArrayLike = new(StringComparer.Ordinal)
        { "SUMPRODUCT", "MMULT", "TRANSPOSE", "DSUM", "DCOUNT", "DGET", "DAVERAGE" };
    private static readonly HashSet<string> Dynamic = new(StringComparer.Ordinal)
        { "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY" };
    private static readonly HashSet<string> LambdaFns = new(StringComparer.Ordinal)
        { "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY" };
    private static readonly HashSet<string> Cheap = new(StringComparer.Ordinal)
        { "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTIF", "COUNTIFS", "AVERAGE", "AVERAGEIF",
          "AVERAGEIFS", "MAX", "MIN", "MAXIFS", "MINIFS", "IF", "IFERROR", "ROUND", "ABS",
          "AND", "OR", "CONCAT", "TEXT" };

    [GeneratedRegex(@"_xlfn\.(_xlws\.)?|_xlpm\.", RegexOptions.IgnoreCase)]
    private static partial Regex PrefixRe();

    [GeneratedRegex(@"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")]
    private static partial Regex FuncRe();

    // Full-column ref (A:A / $A:$H), including sheet-qualified (Data!$A:$H).
    [GeneratedRegex(@"(?<![A-Za-z0-9_$])\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![A-Za-z0-9_])")]
    private static partial Regex FullColRe();

    // A bounded A1 range like $C$2:$C$10001 (possibly sheet-qualified — match anywhere).
    [GeneratedRegex(@"\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)")]
    private static partial Regex RangeRe();

    // An A1 cell token. The cell part of a sheet-qualified ref (Data!F2) IS normalised
    // too (the sheet prefix stays), so copied rows collapse to one pattern — matching the
    // Python oracle. (Not preceded by letter/digit/_/$, and not a function call.)
    [GeneratedRegex(@"(?<![A-Za-z0-9_$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")]
    private static partial Regex CellRe();

    /// <summary>Strip a leading '=' and Excel's internal function prefixes.</summary>
    public static string Clean(string? formula)
    {
        var f = formula ?? "";
        if (f.StartsWith('=')) f = f[1..];
        return PrefixRe().Replace(f, "");
    }

    /// <summary>All function names used, uppercased, prefixes removed.</summary>
    public static List<string> Functions(string? formula)
    {
        var clean = Clean(formula);
        var result = new List<string>();
        foreach (Match m in FuncRe().Matches(clean))
            result.Add(m.Groups[1].Value.ToUpperInvariant());
        return result;
    }

    public static FunctionClass Classify(string? formula)
    {
        var fns = new HashSet<string>(Functions(formula), StringComparer.Ordinal);
        if (fns.Overlaps(Volatile)) return FunctionClass.Volatile;
        if (fns.Overlaps(ArrayLike)) return FunctionClass.Array;
        if (fns.Overlaps(LambdaFns)) return FunctionClass.Lambda;
        if (fns.Overlaps(Dynamic)) return FunctionClass.DynamicArray;
        if (fns.Overlaps(Lookup)) return FunctionClass.Lookup;
        if (fns.Contains("LET") && !DropLet(fns).Overlaps(Volatile) &&
            !DropLet(fns).Overlaps(ArrayLike) && !DropLet(fns).Overlaps(Lookup))
            return FunctionClass.Let;
        if (fns.Overlaps(Cheap)) return FunctionClass.Cheap;
        return fns.Count > 0 ? FunctionClass.Other : FunctionClass.Cheap;

        static HashSet<string> DropLet(HashSet<string> s)
        {
            var c = new HashSet<string>(s, StringComparer.Ordinal);
            c.Remove("LET");
            return c;
        }
    }

    public static bool IsVolatile(string? formula) =>
        new HashSet<string>(Functions(formula), StringComparer.Ordinal).Overlaps(Volatile);

    public static bool IsSingleThreaded(string? formula) =>
        new HashSet<string>(Functions(formula), StringComparer.Ordinal).Overlaps(SingleThreaded);

    public static bool HasFullColumn(string? formula) => FullColRe().IsMatch(Clean(formula));

    public static int ColToNum(string col)
    {
        var n = 0;
        foreach (var ch in col.ToUpperInvariant())
            n = n * 26 + (ch - 'A' + 1);
        return n;
    }

    /// <summary>
    /// Estimate referenced cells summed across the formula's ranges. Full-column refs
    /// get a whole-column penalty; bounded ranges are charged by area. Coarse but
    /// monotonic — enough for ranking and bucketing.
    /// </summary>
    public static long CellsTouched(string? formula)
    {
        var f = Clean(formula);
        long total = 0;
        if (HasFullColumn(f))
            total += 1_048_576L * (1 + CountChar(f, ':'));
        foreach (Match m in RangeRe().Matches(f))
        {
            long w = Math.Abs(ColToNum(m.Groups[3].Value) - ColToNum(m.Groups[1].Value)) + 1;
            long h = Math.Abs(long.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) -
                              long.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)) + 1;
            total += w * h;
        }
        return Math.Max(total, 1);
    }

    /// <summary>Bucket cells-touched into log-ish bands so similar-cost cells group.</summary>
    public static int CellsBucket(long n)
    {
        if (n <= 1) return 0;
        if (n >= 1_048_576) return 99;
        return (int)(Math.Log10(n) * 2);
    }

    /// <summary>
    /// Convert an A1 formula at 1-based (row, col) to an R1C1-normalised form so copied
    /// relative formulas collapse to one pattern. Sheet-qualified refs are left fixed.
    /// </summary>
    public static string NormalizeR1C1(string? formula, int row, int col)
    {
        var f = Clean(formula);
        return CellRe().Replace(f, m =>
        {
            var colAbs = m.Groups[1].Value == "$";
            var rowAbs = m.Groups[3].Value == "$";
            var tcol = ColToNum(m.Groups[2].Value);
            var trow = int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
            var c = colAbs ? $"C{tcol}" : $"C[{tcol - col}]";
            var r = rowAbs ? $"R{trow}" : $"R[{trow - row}]";
            return r + c;
        });
    }

    /// <summary>The cost-attribution unit: (R1C1 pattern, cells-touched bucket).</summary>
    public static (string R1C1, int Bucket) GroupingKey(string? formula, int row, int col) =>
        (NormalizeR1C1(formula, row, col), CellsBucket(CellsTouched(formula)));

    private static int CountChar(string s, char c)
    {
        var n = 0;
        foreach (var ch in s) if (ch == c) n++;
        return n;
    }
}
