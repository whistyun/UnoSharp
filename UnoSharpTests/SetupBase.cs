using NUnit.Framework;
using System.IO;
using System.Reflection;

namespace UnoSharp.Tests
{
    public class SetupBase
    {
        [SetUp]
        public void Setup()
        {
            Assembly myAssembly = typeof(SetupBase).Assembly;
            string path = myAssembly.Location;
            Directory.SetCurrentDirectory(Path.GetDirectoryName(path));
        }
    }
}
