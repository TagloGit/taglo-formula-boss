using System.Globalization;
using System.Security.Cryptography;
using System.Text;

using FormulaBoss.Interception;
using FormulaBoss.Parsing;

namespace FormulaBoss.Transpilation;

/// <summary>
///     Transpiles DSL AST to C# source code for UDF generation.
/// </summary>
public class CSharpTranspiler
{
    private readonly HashSet<string> _lambdaParameters = [];

    private readonly Dictionary<string, string>
        _parameterTypes = new(); // param name -> type (e.g., "Cell", "Interior")

    private readonly HashSet<string> _rowParameters = []; // Lambda parameters that are row objects (object[])
    private readonly HashSet<string> _usedColumnBindings = []; // Column bindings that were actually referenced
    private Dictionary<string, ColumnBindingInfo>? _columnBindings; // LET-bound variable name -> (table, column)
    private bool _needsHeaderContext; // True if named column access is used (r[Price], r.Price)
    private bool _requiresObjectModel;
    private bool _usesCols;

    /// <summary>
    ///     Transpiles a DSL expression to a complete C# UDF class.
    /// </summary>
    /// <param name="expression">The parsed AST expression.</param>
    /// <param name="originalSource">The original DSL source text.</param>
    /// <param name="preferredName">Optional preferred name for the UDF (e.g., from a LET variable).</param>
    /// <param name="columnBindings">
    ///     Optional column bindings from LET variables. Maps LET variable names to column binding info
    ///     (e.g., "price" → (tblSales, Price)). Used to resolve r.price or r[price] to column name lookup
    ///     and generate dynamic column parameters.
    /// </param>
    /// <returns>The transpilation result with generated code and metadata.</returns>
    public TranspileResult Transpile(
        Expression expression,
        string originalSource,
        string? preferredName = null,
        Dictionary<string, ColumnBindingInfo>? columnBindings = null)
    {
        _requiresObjectModel = false;
        _usesCols = false;
        _needsHeaderContext = false;
        _columnBindings = columnBindings;
        _lambdaParameters.Clear();
        _parameterTypes.Clear();
        _rowParameters.Clear();
        _usedColumnBindings.Clear();

        // First pass: detect if object model is needed and track rows/cols/map usage
        DetectObjectModelUsage(expression);

        // Generate the method name - use preferred name if provided, otherwise hash
        var methodName = GenerateMethodName(originalSource, preferredName);

        // Generate the expression body
        var expressionCode = TranspileExpression(expression);

        // Generate the complete UDF class
        var sourceCode = GenerateUdfClass(methodName, expressionCode);

        // Return used column bindings so pipeline can build column parameters
        var usedBindings = _usedColumnBindings.Count > 0
            ? _usedColumnBindings.ToList()
            : null;

        return new TranspileResult(sourceCode, methodName, _requiresObjectModel, originalSource, usedBindings);
    }

    private void DetectObjectModelUsage(Expression expression)
    {
        switch (expression)
        {
            case MemberAccess member:
                // .cells triggers object model, .values does not
                // Deep property access (Interior, Font) also requires object model
                if (member.Member.Equals("cells", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("color", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("row", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("col", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("rgb", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("bold", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("italic", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("fontSize", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("format", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("formula", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("address", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("Interior", StringComparison.OrdinalIgnoreCase)
                    || member.Member.Equals("Font", StringComparison.OrdinalIgnoreCase))
                {
                    _requiresObjectModel = true;
                }

                // Track cols usage
                if (member.Member == "cols")
                {
                    _usesCols = true;
                }

                DetectObjectModelUsage(member.Target);
                break;

            case MethodCall call:
                // Track .withHeaders() usage - enables header context
                if (call.Method.Equals("withHeaders", StringComparison.OrdinalIgnoreCase))
                {
                    _needsHeaderContext = true;
                }

                DetectObjectModelUsage(call.Target);
                foreach (var arg in call.Arguments)
                {
                    DetectObjectModelUsage(arg);
                }

                break;

            case BinaryExpr binary:
                DetectObjectModelUsage(binary.Left);
                DetectObjectModelUsage(binary.Right);
                break;

            case UnaryExpr unary:
                DetectObjectModelUsage(unary.Operand);
                break;

            case LambdaExpr lambda:
                DetectObjectModelUsage(lambda.Body);
                break;

            case GroupingExpr grouping:
                DetectObjectModelUsage(grouping.Inner);
                break;

            case IndexAccess indexAccess:
                // Detect named column access: r[Price] where index is an identifier (not a number)
                // This signals we need header context for column name resolution
                if (indexAccess.Index is IdentifierExpr indexIdent
                    && !indexIdent.Name.Equals("null", StringComparison.OrdinalIgnoreCase))
                {
                    _needsHeaderContext = true;
                }

                DetectObjectModelUsage(indexAccess.Target);
                DetectObjectModelUsage(indexAccess.Index);
                break;
        }
    }

    private string TranspileExpression(Expression expression)
    {
        return expression switch
        {
            IdentifierExpr ident => TranspileIdentifier(ident),
            RangeRefExpr => "__source__", // Range references become the input source
            NumberLiteral num => TranspileNumber(num),
            StringLiteral str => TranspileString(str),
            BinaryExpr binary => TranspileBinary(binary),
            UnaryExpr unary => TranspileUnary(unary),
            MemberAccess member => TranspileMemberAccess(member),
            MethodCall call => TranspileMethodCall(call),
            LambdaExpr lambda => TranspileLambda(lambda),
            GroupingExpr grouping => $"({TranspileExpression(grouping.Inner)})",
            IndexAccess indexAccess => TranspileIndexAccess(indexAccess),
            _ => throw new InvalidOperationException($"Unknown expression type: {expression.GetType().Name}")
        };
    }

    private string TranspileIdentifier(IdentifierExpr ident)
    {
        // Handle null keyword
        if (ident.Name.Equals("null", StringComparison.OrdinalIgnoreCase))
        {
            return "null";
        }

        // If it's a lambda parameter, return it directly
        if (_lambdaParameters.Contains(ident.Name))
        {
            return ident.Name;
        }

        // Otherwise it's the input range/data reference
        return "__source__";
    }

    private static string TranspileNumber(NumberLiteral num) => num.Value.ToString(CultureInfo.InvariantCulture);

    private static string TranspileString(StringLiteral str)
    {
        // Escape the string for C#
        var escaped = str.Value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
        return $"\"{escaped}\"";
    }

    private string TranspileBinary(BinaryExpr binary)
    {
        var left = TranspileExpression(binary.Left);
        var right = TranspileExpression(binary.Right);

        // Handle null-coalescing operator - pass through to C#
        if (binary.Operator == "??")
        {
            return $"({left} ?? {right})";
        }

        // Handle null comparisons on member access - wrap in try-catch for COM safety
        if (binary.Operator is "==" or "!=" && _requiresObjectModel)
        {
            var isNullCheck = IsNullLiteral(binary.Right) || IsNullLiteral(binary.Left);
            var isMemberAccessCheck = binary.Left is MemberAccess || binary.Right is MemberAccess;

            if (isNullCheck && isMemberAccessCheck)
            {
                // Wrap null comparisons in try-catch for COM interop safety
                // c.@Comment != null returns false on exception (treating as null)
                // c.@Comment == null returns true on exception (treating as null)
                var defaultValue = binary.Operator == "==" ? "true" : "false";
                return
                    $"((Func<bool>)(() => {{ try {{ return {left} {binary.Operator} {right}; }} catch {{ return {defaultValue}; }} }}))()";
            }
        }

        // For comparison and arithmetic operators in value-only path, lambda parameters and
        // index access on them (e.g., r[0]) are objects and need to be cast to double
        var isComparison = binary.Operator is ">" or "<" or ">=" or "<=" or "==" or "!=";
        var isArithmetic = binary.Operator is "+" or "-" or "*" or "/";
        var isEqualityComparison = binary.Operator is "==" or "!=";

        // Handle string comparisons: when comparing column access to string literal,
        // convert column value to string for proper comparison (object == string uses reference equality)
        if (isEqualityComparison && !_requiresObjectModel)
        {
            // If right side is string literal and left needs cast (column access), use ToString()
            if (IsStringLiteral(binary.Right) && NeedsNumericCast(binary.Left))
            {
                left = $"{left}?.ToString()";
            }

            // If left side is string literal and right needs cast (column access), use ToString()
            if (IsStringLiteral(binary.Left) && NeedsNumericCast(binary.Right))
            {
                right = $"{right}?.ToString()";
            }
        }

        if ((isComparison || isArithmetic) && !_requiresObjectModel)
        {
            // Cast expressions that return objects to double for numeric operations
            if (NeedsNumericCast(binary.Left) && IsNumericLiteral(binary.Right))
            {
                left = WrapInConvertToDouble(left);
            }

            if (NeedsNumericCast(binary.Right) && IsNumericLiteral(binary.Left))
            {
                right = WrapInConvertToDouble(right);
            }

            // Also handle when both sides need casting (e.g., v * v, r[0] * r[1])
            // But NOT for + operator since it can be string concatenation
            var isDefinitelyNumeric = binary.Operator is "-" or "*" or "/";
            if (isDefinitelyNumeric && NeedsNumericCast(binary.Left) && NeedsNumericCast(binary.Right))
            {
                left = WrapInConvertToDouble(left);
                right = WrapInConvertToDouble(right);
            }
            else if (isDefinitelyNumeric)
            {
                // Handle cases with another arithmetic expression (e.g., v * (v + 1), r[0] - r[1])
                // Only for definitely-numeric operators (not + which can be string concatenation)
                if (NeedsNumericCast(binary.Left) && !IsNumericLiteral(binary.Right))
                {
                    left = WrapInConvertToDouble(left);
                }

                if (NeedsNumericCast(binary.Right) && !IsNumericLiteral(binary.Left))
                {
                    right = WrapInConvertToDouble(right);
                }
            }

            // For + operator with row column access: if one side is a lambda parameter (accumulator)
            // and the other is ROW COLUMN ACCESS (index/member on row param), cast the column access.
            // This handles: acc + r[Price], acc + r.Price in reduce operations
            // Only apply for row column access, not general lambda params (to preserve string concatenation)
            if (binary.Operator == "+" && IsRowColumnAccess(binary.Right) && IsLambdaParameter(binary.Left))
            {
                right = WrapInConvertToDouble(right);
            }

            if (binary.Operator == "+" && IsRowColumnAccess(binary.Left) && IsLambdaParameter(binary.Right))
            {
                left = WrapInConvertToDouble(left);
            }
        }

        return $"({left} {binary.Operator} {right})";
    }

    private static bool IsNumericLiteral(Expression expr) => expr is NumberLiteral;

    private static bool IsStringLiteral(Expression expr) => expr is StringLiteral;

    private static bool IsNullLiteral(Expression expr) =>
        expr is IdentifierExpr { Name: "null" };

    private string TranspileUnary(UnaryExpr unary)
    {
        var operand = TranspileExpression(unary.Operand);
        return $"({unary.Operator}{operand})";
    }

    private string TranspileMemberAccess(MemberAccess member)
    {
        var target = TranspileExpression(member.Target);
        var memberName = member.Member;

        // If escaped with @, bypass all validation and pass through verbatim
        if (member.IsEscaped)
        {
            var escapedAccess = $"{target}.{memberName}";
            // If safe access is requested, wrap in try-catch that returns null on exception
            if (member.IsSafeAccess)
            {
                return $"((Func<dynamic>)(() => {{ try {{ return {escapedAccess}; }} catch {{ return null; }} }}))()";
            }

            return escapedAccess;
        }

        // Check for row column access: r.Price where r is a row parameter
        // This takes precedence over Cell property handling when not in object model mode
        if (member.Target is IdentifierExpr targetIdent
            && _rowParameters.Contains(targetIdent.Name)
            && !_requiresObjectModel)
        {
            // Set flag to generate header dictionary in the method
            _needsHeaderContext = true;

            // Check if memberName is a LET-bound column variable (r.price where price = tblSales[Price])
            // In that case, resolve to either a variable reference or the column name
            var columnRef = ResolveColumnBinding(memberName);
            var isVariableRef = columnRef != null && columnRef.EndsWith("_colname_");
            var columnName = columnRef ?? memberName;

            // Row column access via dot notation: r.Price -> r[__GetCol__("Price")] or r[__GetCol__(_price_colname_)]
            var colArg = isVariableRef ? columnName : $"\"{columnName}\"";
            var rowColumnAccess = $"{target}[__GetCol__({colArg})]";
            if (member.IsSafeAccess)
            {
                return $"((Func<dynamic>)(() => {{ try {{ return {rowColumnAccess}; }} catch {{ return null; }} }}))()";
            }

            return rowColumnAccess;
        }

        // Special handling for cell properties on lambda parameters (shorthand syntax)
        if (IsLambdaParameter(member.Target))
        {
            // Handle shorthand properties (color, bold, etc.) - these stay as-is
            var shorthandResult = memberName.ToLowerInvariant() switch
            {
                "value" => _requiresObjectModel ? $"{target}.Value" : target,
                "color" => $"(int)({target}.Interior.ColorIndex ?? 0)",
                "rgb" => $"(int)({target}.Interior.Color ?? 0)",
                "row" => $"{target}.Row",
                "col" => $"{target}.Column",
                "bold" => $"(bool)({target}.Font.Bold ?? false)",
                "italic" => $"(bool)({target}.Font.Italic ?? false)",
                "fontsize" => $"(double)({target}.Font.Size ?? 11)",
                "format" => $"(string)({target}.NumberFormat ?? \"General\")",
                "formula" => $"(string)({target}.Formula ?? \"\")",
                "address" => $"{target}.Address",
                _ => null
            };

            if (shorthandResult != null)
            {
                // Apply safe access wrapper if requested
                if (member.IsSafeAccess)
                {
                    return
                        $"((Func<dynamic>)(() => {{ try {{ return {shorthandResult}; }} catch {{ return null; }} }}))()";
                }

                return shorthandResult;
            }

            // Not a shorthand - check type system for Cell properties
            var paramType = GetExpressionType(member.Target);
            if (paramType != null)
            {
                return TranspileTypedMemberAccess(target, memberName, paramType, member.IsSafeAccess);
            }
        }

        // Special handling for .cells, .values, .rows, .cols on the source
        if (target == "__source__")
        {
            return memberName.ToLowerInvariant() switch
            {
                "cells" => "__cells__",
                "values" => "__values__",
                "rows" => "__rows__",
                "cols" => "__cols__",
                _ => $"__source__.{memberName}"
            };
        }

        // Check if target has a known type (for deep property access like c.Interior.ColorIndex)
        var targetType = GetExpressionType(member.Target);
        if (targetType != null)
        {
            return TranspileTypedMemberAccess(target, memberName, targetType, member.IsSafeAccess);
        }

        // Default: pass through
        var defaultAccess = $"{target}.{memberName}";
        if (member.IsSafeAccess)
        {
            return $"((Func<dynamic>)(() => {{ try {{ return {defaultAccess}; }} catch {{ return null; }} }}))()";
        }

        return defaultAccess;
    }

    /// <summary>
    ///     Transpiles a member access when we know the type of the target.
    ///     Uses the type system to validate properties and apply proper casting.
    /// </summary>
    private string TranspileTypedMemberAccess(string target, string memberName, string targetType,
        bool isSafeAccess = false)
    {
        if (ExcelTypeSystem.Types.TryGetValue(targetType, out var props))
        {
            if (props.TryGetValue(memberName, out var prop))
            {
                // Apply the template with proper casting
                var result = string.Format(prop.Template, target);
                if (isSafeAccess)
                {
                    return $"((Func<dynamic>)(() => {{ try {{ return {result}; }} catch {{ return null; }} }}))()";
                }

                return result;
            }

            // Unknown property - throw with suggestion if available
            var similar = ExcelTypeSystem.FindSimilar(targetType, memberName);
            if (similar != null)
            {
                throw new TranspileException(
                    $"Unknown property '{memberName}' on {targetType}. Did you mean '{similar}'?");
            }

            throw new TranspileException($"Unknown property '{memberName}' on {targetType}.");
        }

        // Type not in system - pass through
        var passThrough = $"{target}.{memberName}";
        if (isSafeAccess)
        {
            return $"((Func<dynamic>)(() => {{ try {{ return {passThrough}; }} catch {{ return null; }} }}))()";
        }

        return passThrough;
    }

    /// <summary>
    ///     Gets the type of an expression based on type tracking context.
    /// </summary>
    private string? GetExpressionType(Expression expr)
    {
        return expr switch
        {
            IdentifierExpr id when _parameterTypes.TryGetValue(id.Name, out var t) => t,
            MemberAccess ma => GetMemberAccessResultType(ma),
            _ => null
        };
    }

    /// <summary>
    ///     Gets the result type of a member access expression by looking up in the type system.
    /// </summary>
    private string? GetMemberAccessResultType(MemberAccess ma)
    {
        var targetType = GetExpressionType(ma.Target);
        if (targetType != null && ExcelTypeSystem.Types.TryGetValue(targetType, out var props))
        {
            if (props.TryGetValue(ma.Member, out var prop))
            {
                return prop.ResultType;
            }
        }

        return null;
    }

    private string TranspileMethodCall(MethodCall call)
    {
        // Check if target is a safe-access member - if so, wrap entire chain in try-catch
        if (call.Target is MemberAccess ma && ma.IsSafeAccess)
        {
            // Transpile target WITHOUT safe-access (we'll wrap the whole chain)
            var unsafeTarget = TranspileMemberAccessUnsafe(ma);
            var safeArgs = call.Arguments.Select(TranspileExpression).ToList();
            var methodCall = $"{unsafeTarget}.{call.Method}({string.Join(", ", safeArgs)})";
            return $"((Func<dynamic>)(() => {{ try {{ return {methodCall}; }} catch {{ return null; }} }}))()";
        }

        var target = TranspileExpression(call.Target);
        var args = call.Arguments.Select(TranspileExpression).ToList();

        // Handle .withHeaders() BEFORE the implicit __values__ conversion
        // .withHeaders() should preserve __source__ so .rows can become __rows__
        if (call.Method.Equals("withHeaders", StringComparison.OrdinalIgnoreCase))
        {
            return target; // Pass through, flag is already set during detection
        }

        // If calling a LINQ method directly on the source (without .values or .cells),
        // implicitly use .values - e.g., data.where(...) becomes __values__.Where(...)
        if (target == "__source__")
        {
            target = "__values__";
        }

        // Map DSL methods to C#/LINQ equivalents
        return call.Method.ToLowerInvariant() switch
        {
            "where" => TranspileRowAwareMethod(target, args, call, "Where"),
            "select" => TranspileRowAwareMethod(target, args, call, "Select"),
            "map" => $"__MapPreserveShape__({string.Join(", ", args)})",
            "toarray" => $"{target}.ToArray()",
            "orderby" => $"{target}.OrderBy({string.Join(", ", args)})",
            "orderbydesc" => $"{target}.OrderByDescending({string.Join(", ", args)})",
            // take/skip with negative support: take(-2) gets last 2, skip(-2) skips last 2
            "take" => GenerateTakeSkip(target, args, true),
            "skip" => GenerateTakeSkip(target, args, false),
            "distinct" => $"{target}.Distinct()",
            // For numeric aggregations, cast objects to double
            // On object model path, items might be COM cells - extract .Value if so
            // Use Aggregate() with explicit types to avoid dynamic type inference issues
            // When selector is provided, wrap selector body in Convert.ToDouble()
            "sum" => args.Count > 0
                ? $"{target}.Sum({WrapSelectorInConvertToDouble(args[0])})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(0.0, (acc, x) => acc + Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Sum()",
            "avg" or "average" => args.Count > 0
                ? $"{target}.Average({WrapSelectorInConvertToDouble(args[0])})"
                : _requiresObjectModel
                    ? $"((Func<double>)(() => {{ var items = {target}.ToList(); return items.Count == 0 ? 0.0 : items.Aggregate(0.0, (acc, x) => acc + Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)) / items.Count; }}))()"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Average()",
            "min" => args.Count > 0
                ? $"{target}.Min({WrapSelectorInConvertToDouble(args[0])})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(double.MaxValue, (acc, x) => Math.Min(acc, Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Min()",
            "max" => args.Count > 0
                ? $"{target}.Max({WrapSelectorInConvertToDouble(args[0])})"
                : _requiresObjectModel
                    ? $"{target}.Aggregate(double.MinValue, (acc, x) => Math.Max(acc, Convert.ToDouble(x != null && x.GetType().IsCOMObject ? ((dynamic)x).Value : x)))"
                    : $"{target}.Select(x => Convert.ToDouble(x)).Max()",
            "count" => $"{target}.Count()",
            "first" => $"{target}.First()",
            "firstordefault" => $"{target}.FirstOrDefault()",
            "last" => $"{target}.Last()",
            "lastordefault" => $"{target}.LastOrDefault()",
            // groupBy: with 1 arg, groups and flattens; with 2 args, aggregates per group
            "groupby" => GenerateGroupBy(target, args),
            // aggregate/reduce: custom fold/reduce operation
            "aggregate" or "reduce" => GenerateRowAwareAggregate(target, args, call),
            // scan: running reduction returning array of intermediate states
            "scan" => GenerateRowAwareScan(target, args, call),
            // find: return first matching row, or null if not found
            "find" => TranspileRowAwareMethod(target, args, call, "FirstOrDefault"),
            // some: return true if any row matches predicate
            "some" => TranspileRowAwareMethod(target, args, call, "Any"),
            // every: return true if all rows match predicate
            "every" => TranspileRowAwareMethod(target, args, call, "All"),
            _ => $"{target}.{call.Method}({string.Join(", ", args)})"
        };
    }

    /// <summary>
    ///     Generates take/skip with negative argument support.
    ///     Positive n: Take(n)/Skip(n) - take/skip first n elements
    ///     Negative n: TakeLast(-n)/SkipLast(-n) - take/skip last n elements
    /// </summary>
    private static string GenerateTakeSkip(string target, List<string> args, bool isTake)
    {
        if (args.Count == 0)
        {
            return isTake ? $"{target}.Take(0)" : $"{target}.Skip(0)";
        }

        var arg = args[0];

        // Strip parentheses from unary expressions like "(-2)" to get "-2"
        var trimmedArg = arg;
        while (trimmedArg.StartsWith('(') && trimmedArg.EndsWith(')'))
        {
            trimmedArg = trimmedArg[1..^1];
        }

        // Check if it's a literal negative number (compile-time optimization)
        if (trimmedArg.StartsWith('-') && int.TryParse(trimmedArg, out var literalValue) && literalValue < 0)
        {
            var positiveValue = -literalValue;
            return isTake
                ? $"{target}.TakeLast({positiveValue})"
                : $"{target}.SkipLast({positiveValue})";
        }

        // Check if it's a literal positive number
        if (int.TryParse(trimmedArg, out _))
        {
            return isTake ? $"{target}.Take({trimmedArg})" : $"{target}.Skip({trimmedArg})";
        }

        // Runtime check for variable arguments
        var takeMethod = isTake ? "Take" : "Skip";
        var takeLastMethod = isTake ? "TakeLast" : "SkipLast";
        return $"(({arg}) >= 0 ? {target}.{takeMethod}({arg}) : {target}.{takeLastMethod}(-({arg})))";
    }

    /// <summary>
    ///     Wraps an expression in Convert.ToDouble() for numeric operations.
    /// </summary>
    private static string WrapInConvertToDouble(string expr) => $"Convert.ToDouble({expr})";

    /// <summary>
    ///     Wraps a lambda selector's body in Convert.ToDouble() for numeric aggregations.
    ///     e.g., "r => r[1]" becomes "r => Convert.ToDouble(r[1])"
    /// </summary>
    private static string WrapSelectorInConvertToDouble(string selector)
    {
        var arrowIndex = selector.IndexOf("=>", StringComparison.Ordinal);
        if (arrowIndex > 0)
        {
            var paramPart = selector[..arrowIndex].Trim();
            var bodyPart = selector[(arrowIndex + 2)..].Trim();
            return $"{paramPart} => {WrapInConvertToDouble(bodyPart)}";
        }

        // Not a lambda, just wrap the whole thing
        return WrapInConvertToDouble(selector);
    }

    /// <summary>
    ///     Generates groupBy with optional per-group aggregation.
    ///     - groupBy(keySelector) - groups and flattens (items stay grouped together)
    ///     - groupBy(keySelector, aggregator) - aggregates per group, returns [key, value] pairs
    /// </summary>
    private static string GenerateGroupBy(string target, List<string> args)
    {
        if (args.Count == 0)
        {
            return $"{target}"; // no-op fallback
        }

        if (args.Count == 1)
        {
            // Single argument: just group and flatten (items stay grouped together)
            return $"{target}.GroupBy({args[0]}).SelectMany(g => g)";
        }

        // Two arguments: keySelector and aggregator
        // groupBy(r => r[0], g => g.sum(r => r[1]))
        // Output: [key, aggregatedValue] pairs as object[] for Excel
        var keySelector = args[0];
        var aggregator = args[1];

        // Extract lambda body from aggregator (e.g., "g => g.Sum(...)" -> "g.Sum(...)")
        var arrowIndex = aggregator.IndexOf("=>", StringComparison.Ordinal);
        if (arrowIndex > 0)
        {
            var paramPart = aggregator[..arrowIndex].Trim(); // "g"
            var bodyPart = aggregator[(arrowIndex + 2)..].Trim(); // "g.Sum(...)"
            return
                $"{target}.GroupBy({keySelector}).Select({paramPart} => new object[] {{ {paramPart}.Key, {bodyPart} }})";
        }

        // Fallback if no arrow found (shouldn't happen for valid lambda)
        return $"{target}.GroupBy({keySelector}).Select(g => new object[] {{ g.Key, {aggregator} }})";
    }

    /// <summary>
    ///     Generates aggregate (fold/reduce) with support for both seeded and unseeded forms.
    ///     - aggregate(seed, (acc, x) => ...) - explicit seed value
    ///     - aggregate((acc, x) => ...) - uses first element as seed
    /// </summary>
    private string GenerateAggregate(string target, List<string> args)
    {
        if (args.Count == 0)
        {
            return $"{target}.Aggregate((acc, x) => acc)"; // no-op fallback
        }

        if (args.Count == 1)
        {
            // Single argument: the accumulator function, no seed
            // aggregate((acc, x) => acc + x) uses first element as seed
            return $"{target}.Aggregate({args[0]})";
        }

        // Two arguments: seed and accumulator function
        // aggregate(0, (acc, x) => acc + x)
        // Convert integer seeds to double to avoid type mismatch when lambda returns double
        var seed = ConvertSeedToDouble(args[0]);
        return $"{target}.Aggregate({seed}, {args[1]})";
    }

    /// <summary>
    ///     Transpiles a method call that may operate on rows, setting row context for lambda parameters.
    ///     This enables named column access like r[Price] and r.Price.
    /// </summary>
    private string TranspileRowAwareMethod(string target, List<string> args, MethodCall call, string csharpMethod)
    {
        // Check if operating on __rows__ - if so, set row context for lambda parameter
        var isRowContext = target == "__rows__" || target.Contains("__rows__");

        if (isRowContext && call.Arguments.Count > 0 && call.Arguments[0] is LambdaExpr lambda)
        {
            // Add the lambda's first parameter to row parameters set
            var rowParam = lambda.Parameters.FirstOrDefault();
            if (rowParam != null)
            {
                _rowParameters.Add(rowParam);
            }

            // Transpile the lambda with row context
            var transpiledLambda = TranspileExpression(lambda);

            // Remove from row parameters
            if (rowParam != null)
            {
                _rowParameters.Remove(rowParam);
            }

            return $"{target}.{csharpMethod}({transpiledLambda})";
        }

        // Not in row context, use normal args
        return $"{target}.{csharpMethod}({string.Join(", ", args)})";
    }

    /// <summary>
    ///     Generates row-aware aggregate (fold/reduce) that sets row context for the accumulator lambda.
    /// </summary>
    private string GenerateRowAwareAggregate(string target, List<string> preTranspiledArgs, MethodCall call)
    {
        // Check if operating on __rows__
        var isRowContext = target == "__rows__" || target.Contains("__rows__");

        if (preTranspiledArgs.Count == 0)
        {
            return $"{target}.Aggregate((acc, x) => acc)"; // no-op fallback
        }

        // For row-aware aggregate, we need to re-transpile the lambda with row context
        if (isRowContext)
        {
            if (call.Arguments.Count == 1 && call.Arguments[0] is LambdaExpr lambda1)
            {
                // Single argument: unseeded aggregate
                // aggregate((acc, r) => acc + r[Price])
                var rowParam = lambda1.Parameters.Count > 1 ? lambda1.Parameters[1] : null;
                if (rowParam != null)
                {
                    _rowParameters.Add(rowParam);
                }

                var transpiledLambda = TranspileExpression(lambda1);

                if (rowParam != null)
                {
                    _rowParameters.Remove(rowParam);
                }

                return $"{target}.Aggregate({transpiledLambda})";
            }

            if (call.Arguments.Count >= 2 && call.Arguments[1] is LambdaExpr lambda2)
            {
                // Two arguments: seed and accumulator
                // aggregate(0, (acc, r) => acc + r[Price])
                var seed = ConvertSeedToDouble(preTranspiledArgs[0]);
                var rowParam = lambda2.Parameters.Count > 1 ? lambda2.Parameters[1] : null;
                if (rowParam != null)
                {
                    _rowParameters.Add(rowParam);
                }

                var transpiledLambda = TranspileExpression(lambda2);

                if (rowParam != null)
                {
                    _rowParameters.Remove(rowParam);
                }

                return $"{target}.Aggregate({seed}, {transpiledLambda})";
            }
        }

        // Fall back to non-row-aware aggregate
        return GenerateAggregate(target, preTranspiledArgs);
    }

    /// <summary>
    ///     Generates row-aware scan (running reduction) that returns an array of intermediate accumulator values.
    ///     scan(seed, (acc, r) => acc + r[Amount]) returns [acc1, acc2, acc3, ...]
    /// </summary>
    private string GenerateRowAwareScan(string target, List<string> preTranspiledArgs, MethodCall call)
    {
        if (preTranspiledArgs.Count < 2)
        {
            return $"{target}"; // Need both seed and lambda
        }

        var seed = ConvertSeedToDouble(preTranspiledArgs[0]);

        // Check if operating on __rows__ with row context
        var isRowContext = target == "__rows__" || target.Contains("__rows__");

        if (isRowContext && call.Arguments.Count >= 2 && call.Arguments[1] is LambdaExpr lambda)
        {
            var rowParam = lambda.Parameters.Count > 1 ? lambda.Parameters[1] : null;
            var accParam = lambda.Parameters.Count > 0 ? lambda.Parameters[0] : "acc";

            if (rowParam != null)
            {
                _rowParameters.Add(rowParam);
            }

            var transpiledLambda = TranspileExpression(lambda);

            if (rowParam != null)
            {
                _rowParameters.Remove(rowParam);
            }

            // Generate scan using Aggregate that collects intermediate values
            // Wrap in a function that builds a list of all intermediate values
            return $"((Func<IEnumerable<object>>)(() => {{ " +
                   $"Func<double, object[], double> fn = {transpiledLambda}; " +
                   $"var results = new List<object>(); " +
                   $"var {accParam} = {seed}; " +
                   $"foreach (var {rowParam} in {target}) {{ " +
                   $"{accParam} = fn({accParam}, {rowParam}); " +
                   $"results.Add({accParam}); " +
                   $"}} " +
                   $"return results; " +
                   $"}}))()";
        }

        // Fall back to simple scan without row context
        var simpleLambda = preTranspiledArgs[1];
        return $"((Func<IEnumerable<object>>)(() => {{ " +
               $"Func<double, object, double> fn = {simpleLambda}; " +
               $"var results = new List<object>(); " +
               $"var acc = {seed}; " +
               $"foreach (var x in {target}) {{ " +
               $"acc = fn(acc, x); " +
               $"results.Add(acc); " +
               $"}} " +
               $"return results; " +
               $"}}))()";
    }

    /// <summary>
    ///     Converts integer literal seeds to double literals for aggregate operations.
    ///     This prevents type mismatches when the accumulator lambda returns double
    ///     (common with Excel values which are often doubles).
    /// </summary>
    private static string ConvertSeedToDouble(string seed)
    {
        // If it's already a double literal (contains . or d/D suffix), leave it alone
        if (seed.Contains('.') || seed.EndsWith("d", StringComparison.OrdinalIgnoreCase))
        {
            return seed;
        }

        // If it's an integer literal, convert to double
        if (int.TryParse(seed, out _))
        {
            return $"{seed}d";
        }

        // Handle parenthesized negative integers like "(-1)"
        if (seed.StartsWith("(-") && seed.EndsWith(")"))
        {
            var inner = seed[2..^1]; // Extract "1" from "(-1)"
            if (int.TryParse(inner, out _))
            {
                return $"(-{inner}d)";
            }
        }

        // Otherwise (could be a variable or expression), leave it alone
        return seed;
    }

    private string TranspileLambda(LambdaExpr lambda)
    {
        // Add all parameters to the tracking set
        foreach (var param in lambda.Parameters)
        {
            _lambdaParameters.Add(param);

            // In object model mode, single-parameter lambdas on .cells iterate over Cell objects
            if (_requiresObjectModel && lambda.Parameters.Count == 1)
            {
                _parameterTypes[param] = "Cell";
            }
        }

        var body = TranspileExpression(lambda.Body);

        // Remove all parameters from the tracking set
        foreach (var param in lambda.Parameters)
        {
            _lambdaParameters.Remove(param);
            _parameterTypes.Remove(param);
        }

        // Generate the lambda syntax
        if (lambda.Parameters.Count == 1)
        {
            return $"{lambda.Parameters[0]} => {body}";
        }

        // Multi-parameter: (acc, x) => body
        return $"({string.Join(", ", lambda.Parameters)}) => {body}";
    }

    private bool IsLambdaParameter(Expression expression) =>
        expression is IdentifierExpr ident && _lambdaParameters.Contains(ident.Name);

    /// <summary>
    ///     Checks if an expression needs numeric casting for comparisons/arithmetic.
    ///     This includes lambda parameters (v, r, c), index access on lambda parameters (r[0], r[Price]),
    ///     and member access on row parameters (r.Price).
    /// </summary>
    private bool NeedsNumericCast(Expression expression)
    {
        return expression switch
        {
            IdentifierExpr ident => _lambdaParameters.Contains(ident.Name),
            IndexAccess indexAccess => NeedsNumericCast(indexAccess
                .Target), // r[0], r[Price] needs cast if r is lambda param
            // r.Price needs cast if r is a row parameter (not in object model mode)
            MemberAccess member when !_requiresObjectModel && member.Target is IdentifierExpr target
                                                           && _rowParameters.Contains(target.Name) => true,
            _ => false
        };
    }

    /// <summary>
    ///     Checks if an expression is row column access (r[Price] or r.Price where r is a row parameter).
    ///     Used to determine when + operator should use numeric casting vs string concatenation.
    /// </summary>
    private bool IsRowColumnAccess(Expression expression)
    {
        return expression switch
        {
            // r[Price] where r is a row parameter and Price is an identifier (named column)
            IndexAccess { Target: IdentifierExpr target, Index: IdentifierExpr } when
                _rowParameters.Contains(target.Name) => true,
            // r.Price where r is a row parameter
            MemberAccess { Target: IdentifierExpr target } when !_requiresObjectModel &&
                                                                _rowParameters.Contains(target.Name) => true,
            _ => false
        };
    }

    private string TranspileIndexAccess(IndexAccess indexAccess)
    {
        var target = TranspileExpression(indexAccess.Target);

        // Check if this is named column access on a row parameter: r[Price]
        // If target is a row parameter and index is an identifier, generate header lookup
        if (indexAccess.Target is IdentifierExpr targetIdent
            && _rowParameters.Contains(targetIdent.Name)
            && indexAccess.Index is IdentifierExpr indexIdent
            && !indexIdent.Name.Equals("null", StringComparison.OrdinalIgnoreCase))
        {
            // Set flag to generate header dictionary in the method
            _needsHeaderContext = true;

            // Check if identifier is a LET-bound column variable (r[price] where price = tblSales[Price])
            // In that case, resolve to either a variable reference or the column name
            var columnRef = ResolveColumnBinding(indexIdent.Name);
            var isVariableRef = columnRef != null && columnRef.EndsWith("_colname_");
            var columnName = columnRef ?? indexIdent.Name;

            // Named column access: r[Price] -> r[__GetCol__("Price")] or r[__GetCol__(_price_colname_)]
            var colArg = isVariableRef ? columnName : $"\"{columnName}\"";
            return $"{target}[__GetCol__({colArg})]";
        }

        // Handle negative index on row parameter: r[-1] accesses last column
        if (indexAccess.Target is IdentifierExpr rowIdent
            && _rowParameters.Contains(rowIdent.Name)
            && indexAccess.Index is UnaryExpr { Operator: "-", Operand: NumberLiteral negNum })
        {
            // r[-1] -> r[r.Length - 1], r[-2] -> r[r.Length - 2]
            var offset = (int)negNum.Value;
            return $"{target}[{target}.Length - {offset}]";
        }

        var index = TranspileExpression(indexAccess.Index);
        return $"{target}[{index}]";
    }

    /// <summary>
    ///     Resolves a LET-bound column variable to a column name or variable reference.
    ///     For example, if LET has "price, tblSales[Price]", then "price" resolves to either:
    ///     - The literal "Price" if dynamic column params are not enabled
    ///     - A variable reference "_price_colname_" if dynamic column params are enabled
    /// </summary>
    /// <param name="variableName">The variable name to resolve.</param>
    /// <param name="useVariableReference">
    ///     If true, returns a variable reference (like "_price_colname_") for use with dynamic column params.
    ///     If false, returns the literal column name.
    /// </param>
    /// <returns>The column name/variable reference if it's a bound column variable, null otherwise.</returns>
    private string? ResolveColumnBinding(string variableName, bool useVariableReference = true)
    {
        if (_columnBindings != null && _columnBindings.TryGetValue(variableName, out var bindingInfo))
        {
            // Track that this column binding was used
            _usedColumnBindings.Add(variableName);

            if (useVariableReference)
            {
                // Return a variable reference that will be set from the UDF parameter
                return $"_{variableName.ToLowerInvariant()}_colname_";
            }

            return bindingInfo.ColumnName;
        }

        return null;
    }

    /// <summary>
    ///     Transpiles a MemberAccess without the safe-access wrapper.
    ///     Used when the safe-access needs to wrap a larger expression (e.g., method call chain).
    /// </summary>
    private string TranspileMemberAccessUnsafe(MemberAccess member)
    {
        var target = TranspileExpression(member.Target);
        var memberName = member.Member;

        // If escaped with @, bypass validation and pass through verbatim (no safe wrapper)
        if (member.IsEscaped)
        {
            return $"{target}.{memberName}";
        }

        // For non-escaped, use same logic as TranspileMemberAccess but without safe wrappers
        // (This is a simplified version - full type system lookup could be added if needed)
        return $"{target}.{memberName}";
    }

    private string GenerateUdfClass(string methodName, string expressionCode)
    {
        var sb = new StringBuilder();

        sb.AppendLine("using System;");
        sb.AppendLine("using System.Linq;");
        sb.AppendLine("using System.Reflection;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine();
        sb.AppendLine("public static class GeneratedUdf");
        sb.AppendLine("{");

        // Get the list of used column bindings for generating extra parameters
        var usedBindings = _usedColumnBindings.ToList();

        if (_requiresObjectModel)
        {
            GenerateObjectModelMethod(sb, methodName, expressionCode, _usesCols, _needsHeaderContext, usedBindings);
        }
        else
        {
            GenerateValueOnlyMethod(sb, methodName, expressionCode, _usesCols, _needsHeaderContext, usedBindings);
        }

        sb.AppendLine("}");

        return sb.ToString();
    }

    private static void GenerateObjectModelMethod(
        StringBuilder sb,
        string methodName,
        string expressionCode,
        bool usesCols,
        bool needsHeaderContext,
        IReadOnlyList<string> usedColumnBindings)
    {
        // Generate the ExcelDNA entry point - uses reflection to avoid assembly identity issues
        // Add extra parameters for each used column binding
        sb.Append("    public static object ").Append(methodName).Append("(object rangeRef");
        foreach (var binding in usedColumnBindings)
        {
            sb.Append(", object ").Append(binding.ToLowerInvariant()).Append("_col_param");
        }

        sb.AppendLine(")");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");

        // Generate column name variable extraction from parameters
        if (usedColumnBindings.Count > 0)
        {
            sb.AppendLine("            // Extract column names from parameters (dynamic column resolution)");
            sb.AppendLine(
                "            // Handle both string values and ExcelReference (if INDEX() didn't force value evaluation)");
            sb.AppendLine("            Func<object, string> extractColName = (param) => {");
            sb.AppendLine("                if (param == null) return \"\";");
            sb.AppendLine("                if (param is string s) return s;");
            sb.AppendLine("                // Check if it's an ExcelReference and extract the actual value");
            sb.AppendLine("                if (param.GetType().Name == \"ExcelReference\")");
            sb.AppendLine("                {");
            sb.AppendLine(
                "                    var getValueMethod = param.GetType().GetMethod(\"GetValue\", Type.EmptyTypes);");
            sb.AppendLine("                    var value = getValueMethod?.Invoke(param, null);");
            sb.AppendLine("                    return value?.ToString() ?? \"\";");
            sb.AppendLine("                }");
            sb.AppendLine("                return param.ToString() ?? \"\";");
            sb.AppendLine("            };");
            sb.AppendLine();
            foreach (var binding in usedColumnBindings)
            {
                var varName = $"_{binding.ToLowerInvariant()}_colname_";
                var paramName = $"{binding.ToLowerInvariant()}_col_param";
                sb.Append("            var ").Append(varName).Append(" = extractColName(").Append(paramName)
                    .AppendLine(");");
            }

            sb.AppendLine();
        }

        sb.AppendLine("            // Get Range via reflection to avoid assembly identity issues");
        sb.AppendLine(
            "            // Object model features (.cells, .color, etc.) require a range reference - they cannot");
        sb.AppendLine(
            "            // work with array output from a previous UDF because cell formatting is not preserved.");
        sb.AppendLine("            if (rangeRef?.GetType()?.Name != \"ExcelReference\")");
        sb.AppendLine(
            "                return \"ERROR: Cell properties require a range reference, not \" + (rangeRef?.GetType()?.Name ?? \"null\") + \". Use a cell range like A1:B10.\";");
        sb.AppendLine();
        sb.AppendLine("            var excelDnaAssembly = rangeRef.GetType().Assembly;");
        sb.AppendLine(
            "            var excelDnaUtilType = excelDnaAssembly.GetType(\"ExcelDna.Integration.ExcelDnaUtil\");");
        sb.AppendLine(
            "            var appProperty = excelDnaUtilType?.GetProperty(\"Application\", BindingFlags.Public | BindingFlags.Static);");
        sb.AppendLine("            dynamic app = appProperty?.GetValue(null);");
        sb.AppendLine("            if (app == null) return \"ERROR: Could not get Excel Application\";");
        sb.AppendLine();
        sb.AppendLine("            // Get range address using ExcelReference's own methods instead of XlCall");
        sb.AppendLine("            var sheetIdProp = rangeRef.GetType().GetProperty(\"SheetId\");");
        sb.AppendLine("            var rowFirstProp = rangeRef.GetType().GetProperty(\"RowFirst\");");
        sb.AppendLine("            var rowLastProp = rangeRef.GetType().GetProperty(\"RowLast\");");
        sb.AppendLine("            var colFirstProp = rangeRef.GetType().GetProperty(\"ColumnFirst\");");
        sb.AppendLine("            var colLastProp = rangeRef.GetType().GetProperty(\"ColumnLast\");");
        sb.AppendLine();
        sb.AppendLine("            var rowFirst = (int)rowFirstProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var rowLast = (int)rowLastProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var colFirst = (int)colFirstProp.GetValue(rangeRef) + 1;");
        sb.AppendLine("            var colLast = (int)colLastProp.GetValue(rangeRef) + 1;");
        sb.AppendLine();
        sb.AppendLine("            // Convert column numbers to letters");
        sb.AppendLine("            Func<int, string> colToLetter = (col) => {");
        sb.AppendLine("                string result = \"\";");
        sb.AppendLine(
            "                while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }");
        sb.AppendLine("                return result;");
        sb.AppendLine("            };");
        sb.AppendLine();
        sb.AppendLine("            string address;");
        sb.AppendLine("            if (rowFirst == rowLast && colFirst == colLast)");
        sb.AppendLine("                address = colToLetter(colFirst) + rowFirst;");
        sb.AppendLine("            else");
        sb.AppendLine(
            "                address = colToLetter(colFirst) + rowFirst + \":\" + colToLetter(colLast) + rowLast;");
        sb.AppendLine();
        sb.AppendLine("            dynamic range = app.Range[address];");
        // Call _Core with column name parameters
        sb.Append("            return ").Append(methodName).Append("_Core(range");
        foreach (var binding in usedColumnBindings)
        {
            var varName = $"_{binding.ToLowerInvariant()}_colname_";
            sb.Append(", ").Append(varName);
        }

        sb.AppendLine(");");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Generate the testable core method (takes Range directly, no ExcelDNA dependencies)
        // Also takes column name parameters for dynamic column resolution
        sb.AppendLine("    /// <summary>");
        sb.AppendLine("    /// Core computation logic - can be called directly with an Excel Range for testing.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static object ").Append(methodName).Append("_Core(dynamic range");
        foreach (var binding in usedColumnBindings)
        {
            sb.Append(", string ").Append("_").Append(binding.ToLowerInvariant()).Append("_colname_");
        }

        sb.AppendLine(")");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            // Build cell list manually - COM collections don't work well with LINQ Cast<T>");
        sb.AppendLine("            var __cells__ = new System.Collections.Generic.List<dynamic>();");
        sb.AppendLine("            foreach (dynamic cell in range.Cells) __cells__.Add(cell);");
        sb.AppendLine();
        sb.AppendLine("            // Get range dimensions");
        sb.AppendLine("            int rowCount = range.Rows.Count;");
        sb.AppendLine("            int colCount = range.Columns.Count;");
        sb.AppendLine();
        // Generate header dictionary if needed for named column access
        if (needsHeaderContext)
        {
            sb.AppendLine("            // Build header dictionary for named column access");
            sb.AppendLine(
                "            // First try to detect if this is an Excel Table (ListObject) and use its headers");
            sb.AppendLine(
                "            var __headers__ = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);");
            sb.AppendLine("            var __skipHeaderRow__ = false;");
            sb.AppendLine("            try");
            sb.AppendLine("            {");
            sb.AppendLine("                dynamic app = range.Application;");
            sb.AppendLine("                dynamic sheet = range.Worksheet;");
            sb.AppendLine("                foreach (dynamic lo in sheet.ListObjects)");
            sb.AppendLine("                {");
            sb.AppendLine("                    // Check if range intersects with this table");
            sb.AppendLine("                    dynamic intersection = app.Intersect(range, lo.Range);");
            sb.AppendLine("                    if (intersection != null)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        // Found a table - use its header row for column names");
            sb.AppendLine("                        dynamic headerRange = lo.HeaderRowRange;");
            sb.AppendLine("                        if (headerRange != null)");
            sb.AppendLine("                        {");
            sb.AppendLine("                            int headerColCount = headerRange.Columns.Count;");
            sb.AppendLine("                            for (int c = 1; c <= headerColCount; c++)");
            sb.AppendLine("                            {");
            sb.AppendLine(
                "                                var headerValue = headerRange.Cells[1, c].Value?.ToString() ?? \"\";");
            sb.AppendLine(
                "                                if (!string.IsNullOrEmpty(headerValue) && !__headers__.ContainsKey(headerValue))");
            sb.AppendLine(
                "                                    __headers__[headerValue] = c - 1;  // Store 0-based index for array access");
            sb.AppendLine("                            }");
            sb.AppendLine("                            // Check if input range includes header row - if so, skip it");
            sb.AppendLine("                            dynamic headerIntersect = app.Intersect(range, headerRange);");
            sb.AppendLine("                            __skipHeaderRow__ = headerIntersect != null;");
            sb.AppendLine("                        }");
            sb.AppendLine("                        break;");
            sb.AppendLine("                    }");
            sb.AppendLine("                }");
            sb.AppendLine("            }");
            sb.AppendLine("            catch { /* ListObject detection failed, fall back to first-row headers */ }");
            sb.AppendLine();
            sb.AppendLine("            // Fall back to first-row headers if no table detected");
            sb.AppendLine("            if (__headers__.Count == 0)");
            sb.AppendLine("            {");
            sb.AppendLine("                for (int c = 1; c <= colCount; c++)");
            sb.AppendLine("                {");
            sb.AppendLine("                    var headerValue = range.Cells[1, c].Value?.ToString() ?? \"\";");
            sb.AppendLine(
                "                    if (!string.IsNullOrEmpty(headerValue) && !__headers__.ContainsKey(headerValue))");
            sb.AppendLine(
                "                        __headers__[headerValue] = c - 1;  // Store 0-based index for array access");
            sb.AppendLine("                }");
            sb.AppendLine("                __skipHeaderRow__ = true;  // First row is headers, skip it");
            sb.AppendLine("            }");
            sb.AppendLine();
            sb.AppendLine("            // Column lookup helper with detailed error message");
            sb.AppendLine("            Func<string, int> __GetCol__ = (name) => {");
            sb.AppendLine("                if (__headers__.TryGetValue(name, out var idx)) return idx;");
            sb.AppendLine(
                "                throw new Exception($\"Column '{name}' not found. Available columns: {string.Join(\", \", __headers__.Keys)}\");");
            sb.AppendLine("            };");
            sb.AppendLine();
            sb.AppendLine("            // Build row arrays for .rows operations - skip header row if present in range");
            sb.AppendLine("            var __dataStartRow__ = __skipHeaderRow__ ? 2 : 1;");
            sb.AppendLine("            var __dataRowCount__ = __skipHeaderRow__ ? rowCount - 1 : rowCount;");
            sb.AppendLine(
                "            var __rows__ = Enumerable.Range(__dataStartRow__, __dataRowCount__).Select(r =>");
            sb.AppendLine(
                "                Enumerable.Range(1, colCount).Select(c => (object)range.Cells[r, c].Value).ToArray());");
        }
        else
        {
            sb.AppendLine("            // Build row arrays for .rows operations (each row is an object[] of values)");
            sb.AppendLine("            var __rows__ = Enumerable.Range(1, rowCount).Select(r =>");
            sb.AppendLine(
                "                Enumerable.Range(1, colCount).Select(c => (object)range.Cells[r, c].Value).ToArray());");
        }

        sb.AppendLine();
        sb.AppendLine("            // Build column arrays for .cols operations");
        sb.AppendLine("            var __cols__ = Enumerable.Range(1, colCount).Select(c =>");
        sb.AppendLine(
            "                Enumerable.Range(1, rowCount).Select(r => (object)range.Cells[r, c].Value).ToArray());");
        sb.AppendLine();
        sb.AppendLine("            // Helper for .map() - preserves 2D shape");
        sb.AppendLine("            Func<Func<dynamic, object>, object[,]> __MapPreserveShape__ = (transform) => {");
        sb.AppendLine("                var result = new object[rowCount, colCount];");
        sb.AppendLine("                for (int r = 0; r < rowCount; r++)");
        sb.AppendLine("                    for (int c = 0; c < colCount; c++)");
        sb.AppendLine("                        result[r, c] = transform(range.Cells[r + 1, c + 1]);");
        sb.AppendLine("                return result;");
        sb.AppendLine("            };");
        sb.AppendLine();

        // Replace placeholders and generate the result
        var code = expressionCode
            .Replace("__source__", "range")
            .Replace("__cells__", "__cells__")
            .Replace("__values__", "range.Value");

        sb.Append("            object result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            // Normalize result inline");
        sb.AppendLine("            if (result == null) return string.Empty;");
        sb.AppendLine(
            "            if (result is string || result is double || result is int || result is bool) return result;");
        sb.AppendLine("            if (result is object[,]) return result;");
        sb.AppendLine();
        sb.AppendLine("            // Handle IEnumerable<object[]> from .rows or .cols operations");
        sb.AppendLine("            if (result is System.Collections.IEnumerable enumerable && !(result is string))");
        sb.AppendLine("            {");
        sb.AppendLine("                var list = enumerable.Cast<object>().ToList();");
        sb.AppendLine("                if (list.Count == 0) return string.Empty;");
        sb.AppendLine();
        sb.AppendLine("                // Check if items are arrays (from .rows or .cols)");
        sb.AppendLine("                if (list[0] is object[] firstArray)");
        sb.AppendLine("                {");
        sb.AppendLine("                    var arrayList = list.Cast<object[]>().ToList();");
        if (usesCols)
        {
            // For .cols: each array is a column, output as columns (transposed)
            sb.AppendLine("                    // .cols output: each array becomes a column");
            sb.AppendLine("                    var resultRowCount = arrayList.Max(arr => arr.Length);");
            sb.AppendLine("                    var output = new object[resultRowCount, arrayList.Count];");
            sb.AppendLine("                    for (int col = 0; col < arrayList.Count; col++)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        for (int row = 0; row < arrayList[col].Length; row++)");
            sb.AppendLine("                            output[row, col] = arrayList[col][row] ?? string.Empty;");
            sb.AppendLine("                        // Pad with empty if ragged");
            sb.AppendLine("                        for (int row = arrayList[col].Length; row < resultRowCount; row++)");
            sb.AppendLine("                            output[row, col] = string.Empty;");
            sb.AppendLine("                    }");
        }
        else
        {
            // For .rows: each array is a row, output as rows (default)
            sb.AppendLine("                    // .rows output: each array becomes a row");
            sb.AppendLine("                    var resultColCount = arrayList.Max(arr => arr.Length);");
            sb.AppendLine("                    var output = new object[arrayList.Count, resultColCount];");
            sb.AppendLine("                    for (int i = 0; i < arrayList.Count; i++)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        for (int j = 0; j < arrayList[i].Length; j++)");
            sb.AppendLine("                            output[i, j] = arrayList[i][j] ?? string.Empty;");
            sb.AppendLine("                        // Pad with empty if ragged");
            sb.AppendLine("                        for (int j = arrayList[i].Length; j < resultColCount; j++)");
            sb.AppendLine("                            output[i, j] = string.Empty;");
            sb.AppendLine("                    }");
        }

        sb.AppendLine("                    return output;");
        sb.AppendLine("                }");
        sb.AppendLine();
        sb.AppendLine("                // Single-column output");
        sb.AppendLine("                var singleColOutput = new object[list.Count, 1];");
        sb.AppendLine("                for (var i = 0; i < list.Count; i++)");
        sb.AppendLine("                {");
        sb.AppendLine("                    var item = list[i];");
        sb.AppendLine("                    // If item is a COM cell object, extract its Value");
        sb.AppendLine("                    if (item != null && item.GetType().IsCOMObject)");
        sb.AppendLine("                        singleColOutput[i, 0] = ((dynamic)item).Value ?? string.Empty;");
        sb.AppendLine("                    else");
        sb.AppendLine("                        singleColOutput[i, 0] = item ?? string.Empty;");
        sb.AppendLine("                }");
        sb.AppendLine("                return singleColOutput;");
        sb.AppendLine("            }");
        sb.AppendLine("            return result;");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
    }

    private static void GenerateValueOnlyMethod(
        StringBuilder sb,
        string methodName,
        string expressionCode,
        bool usesCols,
        bool needsHeaderContext,
        IReadOnlyList<string> usedColumnBindings)
    {
        // Generate the ExcelDNA entry point (uses RuntimeHelpers for ExcelReference → values conversion)
        // Add extra parameters for each used column binding
        sb.Append("    public static object ").Append(methodName).Append("(object rangeRef");
        foreach (var binding in usedColumnBindings)
        {
            sb.Append(", object ").Append(binding.ToLowerInvariant()).Append("_col_param");
        }

        sb.AppendLine(")");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");

        // Generate column name variable extraction from parameters
        if (usedColumnBindings.Count > 0)
        {
            sb.AppendLine("            // Extract column names from parameters (dynamic column resolution)");
            sb.AppendLine(
                "            // Handle both string values and ExcelReference (if INDEX() didn't force value evaluation)");
            sb.AppendLine("            Func<object, string> extractColName = (param) => {");
            sb.AppendLine("                if (param == null) return \"\";");
            sb.AppendLine("                if (param is string s) return s;");
            sb.AppendLine("                // Check if it's an ExcelReference and extract the actual value");
            sb.AppendLine("                if (param.GetType().Name == \"ExcelReference\")");
            sb.AppendLine("                {");
            sb.AppendLine(
                "                    var getValueMethod = param.GetType().GetMethod(\"GetValue\", Type.EmptyTypes);");
            sb.AppendLine("                    var value = getValueMethod?.Invoke(param, null);");
            sb.AppendLine("                    return value?.ToString() ?? \"\";");
            sb.AppendLine("                }");
            sb.AppendLine("                return param.ToString() ?? \"\";");
            sb.AppendLine("            };");
            sb.AppendLine();
            foreach (var binding in usedColumnBindings)
            {
                var varName = $"_{binding.ToLowerInvariant()}_colname_";
                var paramName = $"{binding.ToLowerInvariant()}_col_param";
                sb.Append("            var ").Append(varName).Append(" = extractColName(").Append(paramName)
                    .AppendLine(");");
            }

            sb.AppendLine();
        }

        sb.AppendLine("            // Handle both ExcelReference (from range) and object[,] (from previous UDF)");
        sb.AppendLine("            object[,] values;");
        sb.AppendLine("            object[] externalHeaders = null;"); // For ListObject header detection
        sb.AppendLine("            if (rangeRef?.GetType()?.Name == \"ExcelReference\")");
        sb.AppendLine("            {");
        sb.AppendLine("                // Extract values from ExcelReference via reflection");
        sb.AppendLine(
            "                var getValueMethod = rangeRef.GetType().GetMethod(\"GetValue\", Type.EmptyTypes);");
        sb.AppendLine("                var rawResult = getValueMethod?.Invoke(rangeRef, null);");
        sb.AppendLine("                values = rawResult is object[,] arr ? arr : new object[,] { { rawResult } };");

        // Add ListObject detection for value-only path when headers are needed
        if (needsHeaderContext)
        {
            sb.AppendLine();
            sb.AppendLine("                // Try to detect if this is an Excel Table (ListObject) and get headers");
            sb.AppendLine("                try");
            sb.AppendLine("                {");
            sb.AppendLine("                    var excelDnaAssembly = rangeRef.GetType().Assembly;");
            sb.AppendLine(
                "                    var excelDnaUtilType = excelDnaAssembly.GetType(\"ExcelDna.Integration.ExcelDnaUtil\");");
            sb.AppendLine(
                "                    var appProperty = excelDnaUtilType?.GetProperty(\"Application\", BindingFlags.Public | BindingFlags.Static);");
            sb.AppendLine("                    dynamic app = appProperty?.GetValue(null);");
            sb.AppendLine("                    if (app != null)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        // Get range from ExcelReference");
            sb.AppendLine("                        var rowFirstProp = rangeRef.GetType().GetProperty(\"RowFirst\");");
            sb.AppendLine("                        var rowLastProp = rangeRef.GetType().GetProperty(\"RowLast\");");
            sb.AppendLine(
                "                        var colFirstProp = rangeRef.GetType().GetProperty(\"ColumnFirst\");");
            sb.AppendLine("                        var colLastProp = rangeRef.GetType().GetProperty(\"ColumnLast\");");
            sb.AppendLine("                        var rowFirst = (int)rowFirstProp.GetValue(rangeRef) + 1;");
            sb.AppendLine("                        var rowLast = (int)rowLastProp.GetValue(rangeRef) + 1;");
            sb.AppendLine("                        var colFirst = (int)colFirstProp.GetValue(rangeRef) + 1;");
            sb.AppendLine("                        var colLast = (int)colLastProp.GetValue(rangeRef) + 1;");
            sb.AppendLine("                        Func<int, string> colToLetter = (col) => {");
            sb.AppendLine("                            string result = \"\";");
            sb.AppendLine(
                "                            while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }");
            sb.AppendLine("                            return result;");
            sb.AppendLine("                        };");
            sb.AppendLine("                        string address = rowFirst == rowLast && colFirst == colLast");
            sb.AppendLine("                            ? colToLetter(colFirst) + rowFirst");
            sb.AppendLine(
                "                            : colToLetter(colFirst) + rowFirst + \":\" + colToLetter(colLast) + rowLast;");
            sb.AppendLine("                        dynamic range = app.Range[address];");
            sb.AppendLine("                        dynamic sheet = range.Worksheet;");
            sb.AppendLine("                        foreach (dynamic lo in sheet.ListObjects)");
            sb.AppendLine("                        {");
            sb.AppendLine("                            dynamic intersection = app.Intersect(range, lo.Range);");
            sb.AppendLine("                            if (intersection != null)");
            sb.AppendLine("                            {");
            sb.AppendLine(
                "                                // Found a table - check if headers are NOT in the input range");
            sb.AppendLine("                                dynamic headerRange = lo.HeaderRowRange;");
            sb.AppendLine("                                if (headerRange != null)");
            sb.AppendLine("                                {");
            sb.AppendLine(
                "                                    dynamic headerIntersect = app.Intersect(range, headerRange);");
            sb.AppendLine("                                    if (headerIntersect == null)");
            sb.AppendLine("                                    {");
            sb.AppendLine("                                        // Headers NOT in range - extract them externally");
            sb.AppendLine("                                        int headerColCount = headerRange.Columns.Count;");
            sb.AppendLine("                                        externalHeaders = new object[headerColCount];");
            sb.AppendLine("                                        for (int c = 1; c <= headerColCount; c++)");
            sb.AppendLine(
                "                                            externalHeaders[c - 1] = headerRange.Cells[1, c].Value;");
            sb.AppendLine("                                    }");
            sb.AppendLine("                                }");
            sb.AppendLine("                                break;");
            sb.AppendLine("                            }");
            sb.AppendLine("                        }");
            sb.AppendLine("                    }");
            sb.AppendLine("                }");
            sb.AppendLine("                catch { /* ListObject detection failed, use first row as headers */ }");
        }

        sb.AppendLine("            }");
        sb.AppendLine("            else if (rangeRef is object[,] inputArray)");
        sb.AppendLine("            {");
        sb.AppendLine("                // Already an array (from previous UDF in LET chain)");
        sb.AppendLine("                values = inputArray;");
        sb.AppendLine("            }");
        sb.AppendLine("            else if (rangeRef != null)");
        sb.AppendLine("            {");
        sb.AppendLine("                // Single value - wrap in 2D array");
        sb.AppendLine("                values = new object[,] { { rangeRef } };");
        sb.AppendLine("            }");
        sb.AppendLine("            else");
        sb.AppendLine("            {");
        sb.AppendLine("                return \"ERROR: Input is null\";");
        sb.AppendLine("            }");
        sb.AppendLine();

        if (needsHeaderContext)
        {
            // If external headers were found, prepend them to the values array
            sb.AppendLine("            // If external headers from ListObject, prepend to values");
            sb.AppendLine("            if (externalHeaders != null)");
            sb.AppendLine("            {");
            sb.AppendLine("                var rowCount = values.GetLength(0);");
            sb.AppendLine("                var colCount = values.GetLength(1);");
            sb.AppendLine("                var newValues = new object[rowCount + 1, colCount];");
            sb.AppendLine("                for (int c = 0; c < colCount && c < externalHeaders.Length; c++)");
            sb.AppendLine("                    newValues[0, c] = externalHeaders[c];");
            sb.AppendLine("                for (int r = 0; r < rowCount; r++)");
            sb.AppendLine("                    for (int c = 0; c < colCount; c++)");
            sb.AppendLine("                        newValues[r + 1, c] = values[r, c];");
            sb.AppendLine("                values = newValues;");
            sb.AppendLine("            }");
            sb.AppendLine();
        }

        // Call _Core with column name parameters
        sb.Append("            return ").Append(methodName).Append("_Core(values");
        foreach (var binding in usedColumnBindings)
        {
            var varName = $"_{binding.ToLowerInvariant()}_colname_";
            sb.Append(", ").Append(varName);
        }

        sb.AppendLine(");");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // Generate the testable core method (takes values directly, no ExcelDNA dependencies)
        // Also takes column name parameters for dynamic column resolution
        sb.AppendLine("    /// <summary>");
        sb.AppendLine("    /// Core computation logic - can be called directly with a values array for testing.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static object ").Append(methodName).Append("_Core(object[,] values");
        foreach (var binding in usedColumnBindings)
        {
            sb.Append(", string ").Append("_").Append(binding.ToLowerInvariant()).Append("_colname_");
        }

        sb.AppendLine(")");
        sb.AppendLine("    {");
        sb.AppendLine("        try");
        sb.AppendLine("        {");
        sb.AppendLine("            var rowCount = values.GetLength(0);");
        sb.AppendLine("            var colCount = values.GetLength(1);");
        sb.AppendLine("            var __values__ = values.Cast<object>();");
        sb.AppendLine();

        // Generate header dictionary if needed
        if (needsHeaderContext)
        {
            sb.AppendLine("            // Build header dictionary from first row for named column access");
            sb.AppendLine(
                "            var __headers__ = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);");
            sb.AppendLine("            if (rowCount > 0)");
            sb.AppendLine("            {");
            sb.AppendLine("                for (int c = 0; c < colCount; c++)");
            sb.AppendLine("                {");
            sb.AppendLine("                    var headerValue = values[0, c]?.ToString() ?? \"\";");
            sb.AppendLine(
                "                    if (!string.IsNullOrEmpty(headerValue) && !__headers__.ContainsKey(headerValue))");
            sb.AppendLine("                        __headers__[headerValue] = c;");
            sb.AppendLine("                }");
            sb.AppendLine("            }");
            sb.AppendLine();
            sb.AppendLine("            // Column lookup helper with detailed error message");
            sb.AppendLine("            Func<string, int> __GetCol__ = (name) => {");
            sb.AppendLine("                if (__headers__.TryGetValue(name, out var idx)) return idx;");
            sb.AppendLine(
                "                throw new Exception($\"Column '{name}' not found. Available columns: {string.Join(\", \", __headers__.Keys)}\");");
            sb.AppendLine("            };");
            sb.AppendLine();
            sb.AppendLine("            // Build row arrays for .rows operations - skip header row");
            sb.AppendLine("            var __rows__ = Enumerable.Range(1, rowCount - 1)");
            sb.AppendLine(
                "                .Select(r => Enumerable.Range(0, colCount).Select(c => values[r, c]).ToArray());");
        }
        else
        {
            sb.AppendLine("            // Build row arrays for .rows operations");
            sb.AppendLine("            var __rows__ = Enumerable.Range(0, rowCount)");
            sb.AppendLine(
                "                .Select(r => Enumerable.Range(0, colCount).Select(c => values[r, c]).ToArray());");
        }

        sb.AppendLine();
        sb.AppendLine("            // Build column arrays for .cols operations");
        sb.AppendLine("            var __cols__ = Enumerable.Range(0, colCount)");
        sb.AppendLine(
            "                .Select(c => Enumerable.Range(0, rowCount).Select(r => values[r, c]).ToArray());");
        sb.AppendLine();
        sb.AppendLine("            // Helper for .map() - preserves 2D shape");
        sb.AppendLine("            Func<Func<object, object>, object[,]> __MapPreserveShape__ = (transform) => {");
        sb.AppendLine("                var result = new object[rowCount, colCount];");
        sb.AppendLine("                for (int r = 0; r < rowCount; r++)");
        sb.AppendLine("                    for (int c = 0; c < colCount; c++)");
        sb.AppendLine("                        result[r, c] = transform(values[r, c]);");
        sb.AppendLine("                return result;");
        sb.AppendLine("            };");
        sb.AppendLine();

        // Replace placeholders
        var code = expressionCode
            .Replace("__source__", "values")
            .Replace("__values__", "__values__");

        sb.Append("            object result = ").Append(code).AppendLine(";");
        sb.AppendLine();
        sb.AppendLine("            // Normalize result inline");
        sb.AppendLine("            if (result == null) return string.Empty;");
        sb.AppendLine(
            "            if (result is string || result is double || result is int || result is bool) return result;");
        sb.AppendLine("            if (result is object[,]) return result;");
        sb.AppendLine();
        sb.AppendLine("            // Handle IEnumerable<object[]> from .rows or .cols operations");
        sb.AppendLine("            if (result is System.Collections.IEnumerable enumerable && !(result is string))");
        sb.AppendLine("            {");
        sb.AppendLine("                var list = enumerable.Cast<object>().ToList();");
        sb.AppendLine("                if (list.Count == 0) return string.Empty;");
        sb.AppendLine();
        sb.AppendLine("                // Check if items are arrays (from .rows or .cols)");
        sb.AppendLine("                if (list[0] is object[] firstArray)");
        sb.AppendLine("                {");
        sb.AppendLine("                    var arrayList = list.Cast<object[]>().ToList();");
        if (usesCols)
        {
            // For .cols: each array is a column, output as columns (transposed)
            sb.AppendLine("                    // .cols output: each array becomes a column");
            sb.AppendLine("                    var resultRowCount = arrayList.Max(arr => arr.Length);");
            sb.AppendLine("                    var output = new object[resultRowCount, arrayList.Count];");
            sb.AppendLine("                    for (int col = 0; col < arrayList.Count; col++)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        for (int row = 0; row < arrayList[col].Length; row++)");
            sb.AppendLine("                            output[row, col] = arrayList[col][row] ?? string.Empty;");
            sb.AppendLine("                        // Pad with empty if ragged");
            sb.AppendLine("                        for (int row = arrayList[col].Length; row < resultRowCount; row++)");
            sb.AppendLine("                            output[row, col] = string.Empty;");
            sb.AppendLine("                    }");
        }
        else
        {
            // For .rows: each array is a row, output as rows (default)
            sb.AppendLine("                    // .rows output: each array becomes a row");
            sb.AppendLine("                    var resultColCount = arrayList.Max(arr => arr.Length);");
            sb.AppendLine("                    var output = new object[arrayList.Count, resultColCount];");
            sb.AppendLine("                    for (int i = 0; i < arrayList.Count; i++)");
            sb.AppendLine("                    {");
            sb.AppendLine("                        for (int j = 0; j < arrayList[i].Length; j++)");
            sb.AppendLine("                            output[i, j] = arrayList[i][j] ?? string.Empty;");
            sb.AppendLine("                        // Pad with empty if ragged");
            sb.AppendLine("                        for (int j = arrayList[i].Length; j < resultColCount; j++)");
            sb.AppendLine("                            output[i, j] = string.Empty;");
            sb.AppendLine("                    }");
        }

        sb.AppendLine("                    return output;");
        sb.AppendLine("                }");
        sb.AppendLine();
        sb.AppendLine("                // Single-column output");
        sb.AppendLine("                var singleColOutput = new object[list.Count, 1];");
        sb.AppendLine(
            "                for (var i = 0; i < list.Count; i++) singleColOutput[i, 0] = list[i] ?? string.Empty;");
        sb.AppendLine("                return singleColOutput;");
        sb.AppendLine("            }");
        sb.AppendLine("            return result;");
        sb.AppendLine("        }");
        sb.AppendLine("        catch (Exception ex)");
        sb.AppendLine("        {");
        sb.AppendLine("            return \"ERROR: \" + ex.GetType().Name + \": \" + ex.Message;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
    }

    private static string GenerateMethodName(string source, string? preferredName)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            return SanitizeUdfName(preferredName);
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(source));
        var hashString = Convert.ToHexString(hash)[..8].ToLowerInvariant();
        return $"__udf_{hashString}";
    }

    /// <summary>
    ///     Sanitizes a name to be a valid Excel UDF name.
    ///     Converts to uppercase, removes invalid characters, ensures it starts with a letter or underscore.
    /// </summary>
    private static string SanitizeUdfName(string name)
    {
        var upper = name.ToUpperInvariant();
        var sb = new StringBuilder();

        foreach (var c in upper)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
        }

        var result = sb.ToString();

        // Must start with letter or underscore
        if (result.Length == 0)
        {
            return "_UDF";
        }

        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        return result;
    }
}
