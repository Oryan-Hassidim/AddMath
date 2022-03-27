using System;

namespace AddMath
{
    public enum SuggestionType : byte
    {
        Text = 0,
        Matrix = 1
    }
    public class Suggestion : IEquatable<Suggestion>, IComparable<Suggestion>
    {
        public string Text { get; set; }
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
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }

            return left.Equals(right);
        }

        public static bool operator !=(Suggestion left, Suggestion right)
        {
            return !(left == right);
        }

        public static bool operator <(Suggestion left, Suggestion right)
        {
            return ReferenceEquals(left, null) ? !ReferenceEquals(right, null) : left.CompareTo(right) < 0;
        }

        public static bool operator <=(Suggestion left, Suggestion right)
        {
            return ReferenceEquals(left, null) || left.CompareTo(right) <= 0;
        }

        public static bool operator >(Suggestion left, Suggestion right)
        {
            return !ReferenceEquals(left, null) && left.CompareTo(right) > 0;
        }

        public static bool operator >=(Suggestion left, Suggestion right)
        {
            return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.CompareTo(right) >= 0;
        }
    }
}
