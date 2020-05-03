using NUnit.Framework;
using VBScript;

namespace Test
{
    [TestFixture]
    public class VBCompilerTests
    {
        [Test] public void Test()
        {
            var instance = new VBCompiler().Run();
            Assert.IsNotNull(instance);
            Assert.AreEqual("Namespace0.Class0", instance.GetType().FullName);
        }
    }
}
