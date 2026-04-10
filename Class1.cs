using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;

public class EnglishBridgeAddIn : IExcelAddIn
{
    private static Application _xlApp;

    public void AutoOpen()
    {
        // Otteniamo l'istanza dell'applicazione Excel
        _xlApp = (Application)ExcelDnaUtil.Application;

        // 1. SCANSIONE AUTOMATICA (Trova le funzioni storiche come SUM, VLOOKUP, ecc.)
        var scannedFunctions = typeof(WorksheetFunction).GetMethods()
            .Select(m => m.Name.ToUpper())
            .Where(name => !name.StartsWith("GET_") && !name.Equals("TOSTRING"));

        // 2. LA "LISTA INVISIBILE" (Nuove funzioni O365 e comandi logici di base)
        var hiddenFunctions = new string[]
        {
            // Array Dinamici e O365
            "UNIQUE", "FILTER", "SORT", "SORTBY", "XLOOKUP", "XMATCH", "SEQUENCE", "RANDARRAY",
            "BYCOL", "BYROW", "LAMBDA", "LET", "MAKEARRAY", "MAP", "REDUCE", "SCAN",
            // Manipolazione Testo Moderna
            "CONCAT", "TEXTJOIN", "TEXTAFTER", "TEXTBEFORE", "TEXTSPLIT", "VSTACK", "HSTACK",
            // Funzioni Logiche e di Sistema
            "IF", "IFS", "AND", "OR", "NOT", "XOR", "IFERROR", "IFNA", "SWITCH", "CHOOSE",
            "TODAY", "NOW", "ROW", "COLUMN", "ROWS", "COLUMNS", "OFFSET", "INDIRECT"
        };

        // 3. UNIONE DELLE LISTE (Evitando duplicati)
        var allFunctions = scannedFunctions.Concat(hiddenFunctions).Distinct();

        // 4. REGISTRAZIONE DINAMICA DELLE FUNZIONI
        var registrations = allFunctions.Select(name => {
            // Predisponiamo l'accettazione di 10 argomenti (ampiamente sufficienti per l'uso reale)
            Expression<Func<object, object, object, object, object, object, object, object, object, object, object>>
            exp = (p1, p2, p3, p4, p5, p6, p7, p8, p9, p10) => RunBridge(name, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10);

            var reg = new ExcelFunctionRegistration(exp)
            {
                FunctionAttribute = new ExcelFunctionAttribute
                {
                    Name = "EN_" + name,
                    Description = $"Esegue la funzione {name} (Sintassi Inglese/US)",
                    Category = "English Bridge",
                    IsMacroType = true // Permette di leggere gli indirizzi delle celle
                }
            };

            // Istruiamo Excel a passarci i RIFERIMENTI (Indirizzi) invece dei valori nudi
            if (reg.ParameterRegistrations != null)
            {
                foreach (var p in reg.ParameterRegistrations)
                {
                    p.ArgumentAttribute.AllowReference = true;
                }
            }

            return reg;
        });

        registrations.RegisterFunctions();
    }

    /// <summary>
    /// Il motore che esegue effettivamente la formula.
    /// </summary>
    public static object RunBridge(string name, params object[] args)
    {
        try
        {
            // Rimuoviamo gli argomenti "Missing" (quelli che l'utente non ha inserito tra le parentesi)
            var cleanArgs = args.Where(a => !(a is ExcelMissing)).ToList();

            // Formattiamo gli argomenti rimasti rimuovendo i null (es. celle passate ma ignorate dal codice)
            var formatted = cleanArgs.Select(FormatArg).Where(s => s != null).ToList();

            // Costruiamo la stringa in formato US: NOME(Arg1,Arg2...)
            string formula = $"{name}({string.Join(",", formatted)})";

            // Passiamo la palla al motore nativo di Excel
            object result = _xlApp.Evaluate(formula);

            // Se Evaluate fallisce internamente, restituisce un Int32 (codice errore COM)
            if (result is int) return ExcelError.ExcelErrorValue;

            return result;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    /// <summary>
    /// Traduce gli oggetti di Excel in stringhe compatibili con Evaluate.
    /// </summary>
    private static string FormatArg(object arg)
    {
        // 1. Riferimenti: Trasforma la cella nel suo indirizzo (es. Foglio1!$A$1:$B$5)
        if (arg is ExcelReference res)
        {
            object refText = XlCall.Excel(XlCall.xlfReftext, res, true);
            if (refText is string text) return text;
            return null;
        }

        // 2. Numeri: Forza l'uso del punto decimale americano
        if (arg is double d) return d.ToString(CultureInfo.InvariantCulture);

        // 3. Stringhe: Aggiunge le doppie virgolette (es. "Ciao" -> ""Ciao"")
        if (arg is string s) return $"\"{s.Replace("\"", "\"\"")}\"";

        // 4. Booleani: VERO/FALSO
        if (arg is bool b) return b ? "TRUE" : "FALSE";

        // 5. Vuoti: Ignora le celle o gli argomenti vuoti
        if (arg is ExcelEmpty || arg == null) return null;

        // Fallback per tipi non previsti
        return arg.ToString();
    }

    public void AutoClose() { }
}