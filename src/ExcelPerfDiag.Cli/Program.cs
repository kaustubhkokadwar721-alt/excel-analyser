using ExcelPerfDiag.Core.Model;

// P0 scaffold entry point. Real argument parsing + static/full pipeline lands in P1.
Console.WriteLine("ExcelPerfDiag CLI (scaffold).");
Console.WriteLine("Function classes: " + string.Join(", ", Enum.GetValues<FunctionClass>().Select(c => c.ToWire())));
return 0;
