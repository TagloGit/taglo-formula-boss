namespace FormulaBoss.Runtime;

/// <summary>
///     Buffer + delegate bridge for debug-mode trace capture.
///     Generated <c>__FB_&lt;NAME&gt;_DEBUG</c> UDFs call <see cref="Begin" />, <see cref="Set" />,
///     <see cref="Snapshot" />, and <see cref="Return" /> to record per-iteration / per-return
///     snapshots of in-scope locals. <c>FB.LastTrace()</c> reads <see cref="LastBuffer" /> and
///     spills <see cref="TraceBuffer.ToObjectArray" />.
///     <para>
///         Public signatures use only <c>object</c> and primitives so this class is safe to call
///         from Roslyn-generated code loaded into a different assembly context.
///     </para>
/// </summary>
public static class Tracer
{
    /// <summary>Hard cap on snapshot rows per run.</summary>
    public const int MaxRows = 1000;

    private static readonly object _lock = new();
    private static readonly ThreadLocal<TraceBuffer?> _current = new();

    /// <summary>The most recently populated buffer (for <c>FB.LastTrace()</c>).</summary>
    public static TraceBuffer? LastBuffer { get; private set; }

    /// <summary>
    ///     Start a new trace run. Clears any prior buffer for the same caller and makes the
    ///     new buffer current on this thread.
    /// </summary>
    /// <param name="name">The source expression name (e.g. the FB local name).</param>
    /// <param name="callerAddr">The calling cell address (from <c>xlfCaller</c>).</param>
    public static void Begin(string name, string callerAddr)
    {
        var buffer = new TraceBuffer(name, callerAddr);
        lock (_lock)
        {
            _current.Value = buffer;
            LastBuffer = buffer;
        }
    }

    /// <summary>Record a local variable's current value in the live state map.</summary>
    public static void Set(string name, object? value)
    {
        var buffer = _current.Value;
        buffer?.Set(name, value);
    }

    /// <summary>
    ///     Commit a row to the buffer by copying the live state map. <paramref name="kind" /> is
    ///     "entry", "iter", or "return". <paramref name="depth" /> is the loop nesting level
    ///     (0 = outermost). <paramref name="branch" /> is a short label identifying the if/else
    ///     arm taken (may be null).
    /// </summary>
    public static void Snapshot(string kind, int depth, string? branch)
    {
        var buffer = _current.Value;
        buffer?.Snapshot(kind, depth, branch);
    }

    /// <summary>Record the returned value in the "return" column of the live state map.</summary>
    public static void Return(object? value)
    {
        var buffer = _current.Value;
        buffer?.Return(value);
    }

    /// <summary>
    ///     Force a truncation-warning row into the current buffer. Normally this is called
    ///     automatically by <see cref="Snapshot" /> once <see cref="MaxRows" /> is reached, but
    ///     callers may invoke it explicitly.
    /// </summary>
    public static void TruncateWarn()
    {
        var buffer = _current.Value;
        buffer?.TruncateWarn();
    }

    /// <summary>Reset all in-memory state (for tests and add-in reload).</summary>
    public static void Reset()
    {
        lock (_lock)
        {
            _current.Value = null;
            LastBuffer = null;
        }
    }
}

/// <summary>A single trace run's recorded rows plus live variable state.</summary>
public sealed class TraceBuffer
{
    private readonly List<string> _columnOrder = new();
    private readonly HashSet<string> _columnSet = new();
    private readonly Dictionary<string, object?> _liveState = new();
    private readonly object _sync = new();
    private readonly List<Dictionary<string, object?>> _rows = new();
    private int _snapshotCount;
    private bool _truncated;

    public TraceBuffer(string name, string callerAddress)
    {
        Name = name;
        CallerAddress = callerAddress;
    }

    public string Name { get; }
    public string CallerAddress { get; }

    /// <summary>Committed rows, in insertion order. Each row is a snapshot of the live state.</summary>
    public IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows
    {
        get
        {
            lock (_sync)
            {
                return _rows.Select(r => (IReadOnlyDictionary<string, object?>)r).ToList();
            }
        }
    }

    /// <summary>Local variable names in first-seen order (excludes fixed kind/depth/branch columns).</summary>
    public IReadOnlyList<string> LocalNames
    {
        get
        {
            lock (_sync)
            {
                return _columnOrder.ToList();
            }
        }
    }

    internal void Set(string name, object? value)
    {
        lock (_sync)
        {
            _liveState[name] = value;
            if (_columnSet.Add(name))
            {
                _columnOrder.Add(name);
            }
        }
    }

    internal void Snapshot(string kind, int depth, string? branch)
    {
        lock (_sync)
        {
            if (_truncated)
            {
                return;
            }

            if (_snapshotCount >= Tracer.MaxRows)
            {
                AppendTruncationRow();
                return;
            }

            var row = new Dictionary<string, object?>(_liveState)
            {
                ["kind"] = kind,
                ["depth"] = depth,
                ["branch"] = branch ?? string.Empty
            };
            _rows.Add(row);
            _snapshotCount++;
        }
    }

    internal void Return(object? value)
    {
        Set("return", value);
    }

    internal void TruncateWarn()
    {
        lock (_sync)
        {
            if (_truncated)
            {
                return;
            }

            AppendTruncationRow();
        }
    }

    private void AppendTruncationRow()
    {
        _rows.Add(new Dictionary<string, object?>
        {
            ["kind"] = "truncated",
            ["branch"] = $"row cap {Tracer.MaxRows} reached"
        });
        _truncated = true;
    }

    /// <summary>
    ///     Materialise the buffer as a 2D array with a header row. Columns are
    ///     <c>kind, depth, branch</c>, then every local ever set (in first-seen order), with a
    ///     trailing <c>return</c> column if any return value was recorded. Missing cells render
    ///     as the empty string.
    /// </summary>
    public object[,] ToObjectArray()
    {
        lock (_sync)
        {
            var headers = new List<string> { "kind", "depth", "branch" };
            var hasReturn = _columnSet.Contains("return");
            foreach (var name in _columnOrder)
            {
                if (name != "return")
                {
                    headers.Add(name);
                }
            }

            if (hasReturn)
            {
                headers.Add("return");
            }

            var result = new object[_rows.Count + 1, headers.Count];
            for (var c = 0; c < headers.Count; c++)
            {
                result[0, c] = headers[c];
            }

            for (var r = 0; r < _rows.Count; r++)
            {
                var row = _rows[r];
                for (var c = 0; c < headers.Count; c++)
                {
                    result[r + 1, c] = row.TryGetValue(headers[c], out var v) && v != null
                        ? v
                        : string.Empty;
                }
            }

            return result;
        }
    }
}
