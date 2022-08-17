#nullable enable
using System;

namespace AddMath.WordAddIn
{
    public enum SuggestionType : byte
    {
        Text = 0,
        Matrix = 1
    }
    public class Suggestion : IEquatable<Suggestion>, IComparable<Suggestion>
    {
        public string Text { get; set; } = string.Empty;
        public SuggestionType Type { get; set; }

        public override bool Equals(object? obj)
        {
            return obj is Suggestion suggestion &&
                   Text == suggestion.Text &&
                   Type == suggestion.Type;
        }

        public bool Equals(Suggestion? suggestion)
        {
            return Text == suggestion?.Text &&
                   Type == suggestion?.Type;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Text, Type);
        }

        public override string? ToString()
        {
            return $"{Type}: {Text}";
        }

        public int CompareTo(Suggestion? other)
        {
            return Text.CompareTo(other?.Text);
        }

        public static bool operator ==(Suggestion left, Suggestion right)
        {
            if (left is null)
            {
                return right is null;
            }

            return left.Equals(right);
        }

        public static bool operator !=(Suggestion left, Suggestion right)
        {
            return !(left == right);
        }

        public static bool operator <(Suggestion left, Suggestion right)
        {
            return left is null ? right is not null : left.CompareTo(right) < 0;
        }

        public static bool operator <=(Suggestion left, Suggestion right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        public static bool operator >(Suggestion left, Suggestion right)
        {
            return left is not null && left.CompareTo(right) > 0;
        }

        public static bool operator >=(Suggestion left, Suggestion right)
        {
            return left is null ? right is null : left.CompareTo(right) >= 0;
        }
    }
}
