using System.Runtime.InteropServices;

namespace Engine
{
    [ComVisible(true)]
    public interface IEngine
    {
        int Add(int first, int second);
        int First { get; }
        int Second { get; }
        void Record(int result);
    }
}
