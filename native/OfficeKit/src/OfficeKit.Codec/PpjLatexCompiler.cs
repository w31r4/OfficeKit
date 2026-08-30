using OfficeKit.Artifact.Wire.V1;

namespace OfficeKit.Codec;

/// <summary>
/// Finite clean-room LaTeX reader for PPJ formula runs. This is intentionally
/// not TeX: it has no macro expansion, packages, environments, I/O, counters,
/// conditionals, or executable escape hatch.
/// </summary>
internal static class PpjLatexCompiler
{
    private const int MaxSourceLength = 4_096;
    private const int MaxTokens = 512;
    private const int MaxDepth = 32;
    private const int MaxNodes = 2_048;

    private static readonly IReadOnlyDictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["alpha"] = "α", ["beta"] = "β", ["gamma"] = "γ", ["delta"] = "δ", ["epsilon"] = "ε",
        ["varepsilon"] = "ϵ", ["zeta"] = "ζ", ["eta"] = "η", ["theta"] = "θ", ["vartheta"] = "ϑ",
        ["iota"] = "ι", ["kappa"] = "κ", ["lambda"] = "λ", ["mu"] = "μ", ["nu"] = "ν",
        ["xi"] = "ξ", ["pi"] = "π", ["varpi"] = "ϖ", ["rho"] = "ρ", ["varrho"] = "ϱ",
        ["sigma"] = "σ", ["varsigma"] = "ς", ["tau"] = "τ", ["upsilon"] = "υ", ["phi"] = "φ",
        ["varphi"] = "ϕ", ["chi"] = "χ", ["psi"] = "ψ", ["omega"] = "ω",
        ["Gamma"] = "Γ", ["Delta"] = "Δ", ["Theta"] = "Θ", ["Lambda"] = "Λ", ["Xi"] = "Ξ",
        ["Pi"] = "Π", ["Sigma"] = "Σ", ["Upsilon"] = "Υ", ["Phi"] = "Φ", ["Psi"] = "Ψ", ["Omega"] = "Ω",
        ["times"] = "×", ["div"] = "÷", ["cdot"] = "·", ["pm"] = "±", ["mp"] = "∓",
        ["le"] = "≤", ["leq"] = "≤", ["ge"] = "≥", ["geq"] = "≥", ["ne"] = "≠", ["neq"] = "≠",
        ["approx"] = "≈", ["sim"] = "∼", ["equiv"] = "≡", ["propto"] = "∝",
        ["to"] = "→", ["rightarrow"] = "→", ["leftarrow"] = "←", ["leftrightarrow"] = "↔",
        ["in"] = "∈", ["notin"] = "∉", ["subset"] = "⊂", ["subseteq"] = "⊆", ["supset"] = "⊃", ["supseteq"] = "⊇",
        ["cup"] = "∪", ["cap"] = "∩", ["forall"] = "∀", ["exists"] = "∃", ["neg"] = "¬", ["land"] = "∧", ["lor"] = "∨",
        ["infty"] = "∞", ["partial"] = "∂", ["nabla"] = "∇", ["ell"] = "ℓ",
        ["int"] = "∫", ["sum"] = "∑", ["prod"] = "∏",
    };

    private static readonly IReadOnlySet<string> Functions = new HashSet<string>(StringComparer.Ordinal)
    {
        "sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "min", "max", "lim",
    };

    internal static PresentationMathFormula Compile(string source, string? path = null)
    {
        if (string.IsNullOrWhiteSpace(source) || source.Length > MaxSourceLength || source.Any(char.IsControl))
            throw Error("ppj.formula.source", $"PPJ formula source must contain 1 through {MaxSourceLength} printable characters.", path);
        var parser = new Parser(source, path);
        var expression = parser.Parse();
        return new PresentationMathFormula
        {
            SourceLatex = source,
            Expression = expression,
            PlainText = PlainText(expression),
        };
    }

    internal static string PlainText(PresentationMathSequence sequence) => string.Concat(sequence.Nodes.Select(PlainText));

    private static string PlainText(PresentationMathNode node) => node.KindCase switch
    {
        PresentationMathNode.KindOneofCase.Text => node.Text,
        PresentationMathNode.KindOneofCase.Fraction => $"({PlainText(node.Fraction.Numerator)})/({PlainText(node.Fraction.Denominator)})",
        PresentationMathNode.KindOneofCase.Radical => $"√({PlainText(node.Radical.Radicand)})",
        PresentationMathNode.KindOneofCase.Script =>
            PlainText(node.Script.Base) +
            (node.Script.Subscript is null ? string.Empty : $"_({PlainText(node.Script.Subscript)})") +
            (node.Script.Superscript is null ? string.Empty : $"^({PlainText(node.Script.Superscript)})"),
        _ => string.Empty,
    };

    private static CodecException Error(string code, string message, string? path) =>
        path is null ? new CodecException(code, message) : new CodecException(code, message, path);

    private sealed class Parser
    {
        private readonly string source;
        private readonly string? path;
        private int index;
        private int tokens;
        private int nodes;

        internal Parser(string source, string? path)
        {
            this.source = source;
            this.path = path;
        }

        internal PresentationMathSequence Parse()
        {
            var result = ParseSequence(null, 0);
            if (index != source.Length)
                throw Syntax("Formula contains an unexpected closing delimiter.");
            if (result.Nodes.Count == 0)
                throw Syntax("Formula cannot be empty.");
            return result;
        }

        private PresentationMathSequence ParseSequence(char? terminator, int depth)
        {
            if (depth > MaxDepth) throw Budget($"Formula nesting exceeds {MaxDepth} levels.");
            var result = new PresentationMathSequence();
            while (index < source.Length)
            {
                if (terminator is { } end && source[index] == end)
                {
                    index++;
                    return result;
                }
                if (source[index] == '}') throw Syntax("Formula contains an unmatched closing brace.");
                if (char.IsWhiteSpace(source[index]))
                {
                    index++;
                    continue;
                }

                var atom = ParseAtom(depth);
                PresentationMathSequence? subscript = null;
                PresentationMathSequence? superscript = null;
                while (index < source.Length && source[index] is '_' or '^')
                {
                    var marker = source[index++];
                    Token();
                    var argument = ParseScriptArgument(depth + 1);
                    if (argument.Nodes.Count == 0) throw Syntax("Formula script cannot be empty.");
                    if (marker == '_')
                    {
                        if (subscript is not null) throw Syntax("Formula atom has more than one subscript.");
                        subscript = argument;
                    }
                    else
                    {
                        if (superscript is not null) throw Syntax("Formula atom has more than one superscript.");
                        superscript = argument;
                    }
                }

                if (subscript is null && superscript is null)
                {
                    foreach (var node in atom.Nodes) Append(result, node);
                }
                else
                {
                    var script = new PresentationMathScript { Base = atom };
                    if (subscript is not null) script.Subscript = subscript;
                    if (superscript is not null) script.Superscript = superscript;
                    Append(result, Node(script));
                }
            }
            if (terminator is not null) throw Syntax("Formula contains an unclosed group.");
            return result;
        }

        private PresentationMathSequence ParseAtom(int depth)
        {
            if (index >= source.Length) throw Syntax("Formula ended where an atom was required.");
            if (source[index] == '{')
            {
                index++;
                Token();
                return ParseSequence('}', depth + 1);
            }
            if (source[index] == '\\') return ParseCommand(depth);
            if (source[index] is '_' or '^') throw Syntax("Formula script has no base atom.");
            if (source[index] == '%') throw Syntax("Formula comments are not supported; escape a literal percent as \\%.");

            Token();
            var start = index;
            if (char.IsDigit(source[index]))
            {
                while (index < source.Length && (char.IsDigit(source[index]) || source[index] == '.')) index++;
            }
            else
            {
                index++;
            }
            return Sequence(Node(source[start..index]));
        }

        private PresentationMathSequence ParseCommand(int depth)
        {
            index++;
            Token();
            if (index >= source.Length) throw Syntax("Formula ends with an incomplete command.");
            if (!char.IsLetter(source[index]))
            {
                var escaped = source[index++];
                return escaped switch
                {
                    '{' or '}' or '_' or '^' or '%' or '#' or '&' or '$' => Sequence(Node(escaped.ToString())),
                    '\\' => Sequence(Node("\\")),
                    ',' => Sequence(Node(" ")),
                    ';' => Sequence(Node(" ")),
                    ':' => Sequence(Node(" ")),
                    '!' => new PresentationMathSequence(),
                    ' ' => Sequence(Node(" ")),
                    _ => throw Syntax($"Formula command \\{escaped} is not supported."),
                };
            }

            var start = index;
            while (index < source.Length && char.IsLetter(source[index])) index++;
            var command = source[start..index];
            if (Symbols.TryGetValue(command, out var symbol)) return Sequence(Node(symbol));
            if (Functions.Contains(command)) return Sequence(Node(command));
            return command switch
            {
                "frac" => Fraction(depth),
                "sqrt" => Radical(depth),
                "mathrm" or "text" or "operatorname" => RequiredGroup(depth, command),
                "left" or "right" => Delimiter(command),
                "quad" => Sequence(Node(" ")),
                "qquad" => Sequence(Node("  ")),
                _ => throw Error("ppj.formula.unsupportedCommand", $"PPJ formula command \\{command} is outside the supported finite LaTeX subset.", path),
            };
        }

        private PresentationMathSequence Fraction(int depth)
        {
            var numerator = RequiredGroup(depth, "frac numerator");
            var denominator = RequiredGroup(depth, "frac denominator");
            if (numerator.Nodes.Count == 0 || denominator.Nodes.Count == 0)
                throw Syntax("Formula fraction numerator and denominator must be non-empty.");
            return Sequence(Node(new PresentationMathFraction { Numerator = numerator, Denominator = denominator }));
        }

        private PresentationMathSequence Radical(int depth)
        {
            if (index < source.Length && source[index] == '[')
                throw Error("ppj.formula.unsupportedCommand", "Indexed radicals are outside the supported formula subset.", path);
            var radicand = RequiredGroup(depth, "sqrt");
            if (radicand.Nodes.Count == 0) throw Syntax("Formula square root must be non-empty.");
            return Sequence(Node(new PresentationMathRadical { Radicand = radicand }));
        }

        private PresentationMathSequence RequiredGroup(int depth, string command)
        {
            while (index < source.Length && char.IsWhiteSpace(source[index])) index++;
            if (index >= source.Length || source[index] != '{')
                throw Syntax($"Formula {command} requires a braced argument.");
            index++;
            Token();
            return ParseSequence('}', depth + 1);
        }

        private PresentationMathSequence ParseScriptArgument(int depth)
        {
            while (index < source.Length && char.IsWhiteSpace(source[index])) index++;
            if (index < source.Length && source[index] == '{')
            {
                index++;
                Token();
                return ParseSequence('}', depth + 1);
            }
            return ParseAtom(depth);
        }

        private PresentationMathSequence Delimiter(string command)
        {
            while (index < source.Length && char.IsWhiteSpace(source[index])) index++;
            if (index >= source.Length) throw Syntax($"Formula \\{command} requires a delimiter.");
            if (source[index] == '\\')
            {
                index++;
                if (index >= source.Length) throw Syntax($"Formula \\{command} delimiter is incomplete.");
                var escaped = source[index++];
                if (escaped is not ('{' or '}' or '|' or '\\'))
                    throw Syntax($"Formula \\{command} delimiter \\{escaped} is not supported.");
                return Sequence(Node(escaped == '\\' ? "\\" : escaped.ToString()));
            }
            var delimiter = source[index++];
            if (delimiter is not ('(' or ')' or '[' or ']' or '|' or '.'))
                throw Syntax($"Formula \\{command} delimiter {delimiter} is not supported.");
            return delimiter == '.' ? new PresentationMathSequence() : Sequence(Node(delimiter.ToString()));
        }

        private void Append(PresentationMathSequence sequence, PresentationMathNode node)
        {
            if (node.KindCase == PresentationMathNode.KindOneofCase.Text && node.Text.Length == 0) return;
            if (node.KindCase == PresentationMathNode.KindOneofCase.Text &&
                sequence.Nodes.LastOrDefault() is { KindCase: PresentationMathNode.KindOneofCase.Text } previous)
            {
                previous.Text += node.Text;
                return;
            }
            nodes++;
            if (nodes > MaxNodes) throw Budget($"Formula expands beyond {MaxNodes} AST nodes.");
            sequence.Nodes.Add(node);
        }

        private PresentationMathSequence Sequence(PresentationMathNode node)
        {
            var sequence = new PresentationMathSequence();
            Append(sequence, node);
            return sequence;
        }

        private static PresentationMathNode Node(string text) => new() { Text = text };
        private static PresentationMathNode Node(PresentationMathFraction fraction) => new() { Fraction = fraction };
        private static PresentationMathNode Node(PresentationMathRadical radical) => new() { Radical = radical };
        private static PresentationMathNode Node(PresentationMathScript script) => new() { Script = script };

        private void Token()
        {
            tokens++;
            if (tokens > MaxTokens) throw Budget($"Formula exceeds the {MaxTokens}-token budget.");
        }

        private CodecException Syntax(string message) => Error("ppj.formula.syntax", message, path);
        private CodecException Budget(string message) => Error("ppj.formula.budget", message, path);
    }
}
