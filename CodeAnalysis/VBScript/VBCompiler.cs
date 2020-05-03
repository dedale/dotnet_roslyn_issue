using Engine;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.VisualBasic;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;

namespace VBScript
{
    public class BaseClass
    {
        private static void OnForbiddenModalBox(string inputBox, object[] parameters)
        {
            var parameterString = parameters == null ? "<null>" : string.Join(", ", parameters.Select(p => p == null ? "<null>" : $"'{p}'"));
            var msg = $"Call to function '{inputBox}({parameterString})' in VBScript disabled.";
            Console.Error.WriteLine(msg);
            throw new InvalidOperationException(msg);
        }
        public void InputBox(params object[] parameters)
        {
            OnForbiddenModalBox(nameof(InputBox), parameters);
        }
        public void MsgBox(params object[] parameters)
        {
            OnForbiddenModalBox(nameof(MsgBox), parameters);
        }
    }
    public class VBCompiler
    {
        public object Run()
        {
            var code = @"
Option strict Off
Imports Engine
Namespace Namespace0
    Public Class Class0
        Inherits VBScript.BaseClass
        Sub Run(model as IEngine)
            Dim Result As Integer
            Result = model.Add(model.First, model.Second)
            model.Record(Result)
        End Sub
    End Class
End Namespace";
            var tree = VisualBasicSyntaxTree.ParseText(code, VisualBasicParseOptions.Default, encoding: Encoding.UTF8);
            SyntaxTree[] trees = new[] { tree };
            var options = new VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary, optimizationLevel: OptimizationLevel.Release);
            var references =
                new[] { typeof(object), typeof(VBCodeProvider), typeof(IEngine), typeof(BaseClass) }
                .Select(t => MetadataReference.CreateFromFile(WebUtility.UrlDecode(new Uri(t.Assembly.CodeBase).AbsolutePath)))
                .ToList();
            references.ForEach(r => Console.WriteLine(r.FilePath));
            var compilation = VisualBasicCompilation.Create("Script", trees, references, options);
            using (var ms = new MemoryStream())
            {
                var result = compilation.Emit(ms);
                if (!result.Success)
                {
                    var message = new StringBuilder("Compilation failed:\n");
                    result.Diagnostics
                        .Where(d => d.IsWarningAsError || d.Severity == DiagnosticSeverity.Error)
                        .ToList()
                        .ForEach(e => message.AppendLine(e.ToString()));
                    throw new Exception(message.ToString());
                }
                var assembly = Assembly.Load(ms.ToArray());
                return assembly.CreateInstance("Namespace0.Class0");
            }
        }
    }
}
