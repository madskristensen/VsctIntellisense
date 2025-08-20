using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.Imaging;

namespace VsctCompletion.Completion.Providers
{
    public class KnownMonikersList
    {
        // Use Lazy<T> to defer reflection until first access and cache the result
        private static readonly Lazy<IEnumerable<string>> _knownMonikerNames = new(() =>
            typeof(KnownMonikers)
                .GetProperties(BindingFlags.Static | BindingFlags.Public)
                .Select(p => p.Name)
        );

        public static IEnumerable<string> KnownMonikerNames => _knownMonikerNames.Value;
    }
}
