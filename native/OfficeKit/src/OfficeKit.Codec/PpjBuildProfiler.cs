using System.Buffers;
using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OfficeKit.Codec;

/// <summary>
/// Opt-in stage timing for the PPJ build path. The profiler is inactive unless
/// OFFICEKIT_PROFILE_BUILD=1, so ordinary builds do not allocate stage records
/// or write diagnostics to the protocol streams.
/// </summary>
internal sealed class PpjBuildProfiler : IDisposable
{
    private static readonly AsyncLocal<PpjBuildProfiler?> CurrentHolder = new();
    private static readonly IDisposable Noop = new NoopScope();

    private readonly string _operation;
    private readonly long _started;
    private readonly List<Stage> _stages = [];
    private bool _disposed;

    private PpjBuildProfiler(string operation)
    {
        _operation = operation;
        _started = Stopwatch.GetTimestamp();
    }

    internal static PpjBuildProfiler? Start(string operation) =>
        Environment.GetEnvironmentVariable("OFFICEKIT_PROFILE_BUILD") == "1"
            ? new PpjBuildProfiler(operation)
            : null;

    internal static IDisposable Activate(PpjBuildProfiler? profiler)
    {
        var previous = CurrentHolder.Value;
        CurrentHolder.Value = profiler;
        return new Activation(previous);
    }

    internal static IDisposable Measure(string name) =>
        CurrentHolder.Value is { } profiler
            ? profiler.Begin(name)
            : Noop;

    private IDisposable Begin(string name) => new Scope(this, name, Stopwatch.GetTimestamp());

    private void End(string name, long started)
    {
        if (_disposed) return;
        _stages.Add(new Stage(name, Stopwatch.GetElapsedTime(started).TotalMilliseconds));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            var buffer = new ArrayBufferWriter<byte>();
            using (var writer = new Utf8JsonWriter(buffer))
            {
                writer.WriteStartObject();
                writer.WriteString("operation", _operation);
                writer.WriteNumber("totalMilliseconds", Stopwatch.GetElapsedTime(_started).TotalMilliseconds);
                writer.WriteStartArray("stages");
                foreach (var stage in _stages)
                {
                    writer.WriteStartObject();
                    writer.WriteString("name", stage.Name);
                    writer.WriteNumber("milliseconds", stage.Milliseconds);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
                writer.Flush();
            }
            Console.Error.WriteLine($"OFFICEKIT_BUILD_PROFILE {Encoding.UTF8.GetString(buffer.WrittenSpan)}");
        }
        catch (Exception exception) when (exception is IOException or ObjectDisposedException)
        {
            // Profiling must never change the build result when stderr closes.
        }
    }

    private sealed record Stage(string Name, double Milliseconds);

    private sealed class Scope(PpjBuildProfiler profiler, string name, long started) : IDisposable
    {
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            profiler.End(name, started);
        }
    }

    private sealed class Activation(PpjBuildProfiler? previous) : IDisposable
    {
        public void Dispose() => CurrentHolder.Value = previous;
    }

    private sealed class NoopScope : IDisposable
    {
        public void Dispose() { }
    }
}
