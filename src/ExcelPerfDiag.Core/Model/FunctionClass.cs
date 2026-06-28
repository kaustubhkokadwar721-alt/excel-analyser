namespace ExcelPerfDiag.Core.Model;

/// <summary>
/// Function-class taxonomy — the "encoding type" analogue from VertiPaq: it says
/// WHY a pattern costs what it does. Worst-wins ordering when a formula mixes classes.
/// </summary>
public enum FunctionClass
{
    Cheap,
    Lookup,
    Array,
    Volatile,
    DynamicArray,
    Lambda,
    Let,
    Other,
}

public static class FunctionClassExtensions
{
    /// <summary>Stable wire string used in reports and Python-oracle parity diffs.</summary>
    public static string ToWire(this FunctionClass c) => c switch
    {
        FunctionClass.Cheap => "cheap",
        FunctionClass.Lookup => "lookup",
        FunctionClass.Array => "array",
        FunctionClass.Volatile => "volatile",
        FunctionClass.DynamicArray => "dynamic-array",
        FunctionClass.Lambda => "lambda",
        FunctionClass.Let => "let",
        _ => "other",
    };

    public static FunctionClass FromWire(string s) => s switch
    {
        "cheap" => FunctionClass.Cheap,
        "lookup" => FunctionClass.Lookup,
        "array" => FunctionClass.Array,
        "volatile" => FunctionClass.Volatile,
        "dynamic-array" => FunctionClass.DynamicArray,
        "lambda" => FunctionClass.Lambda,
        "let" => FunctionClass.Let,
        _ => FunctionClass.Other,
    };
}
