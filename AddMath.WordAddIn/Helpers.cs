﻿using System.Collections.Generic;
using System.Text;

namespace AddMath.WordAddIn
{
    public static class Helpers
    {
        public static void AppendJoin<T>(this StringBuilder builder, string separator, IEnumerable<T> values)
        {
            foreach (var item in values)
            {
                builder.Append(item);
                builder.Append(separator);
            }
            builder.Remove(builder.Length - separator.Length, separator.Length);
        }
    }
}
