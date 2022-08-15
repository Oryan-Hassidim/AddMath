using System;
using System.Runtime.InteropServices;

namespace AddMath.WordAddIn
{
    [ComVisible(true)]
    public interface IAddMathObject
    {
        void AddMath();
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddMathObject : IAddMathObject, IDisposable
    {
        private Suggestions _suggestions;
        private Suggestions Suggestions => _suggestions ??= new Suggestions();
        public string Theme
        {
            get => _suggestions.Theme;
            set => _suggestions.Theme = value;
        }
        public void AddMath()
        {
            Suggestions.Show();
            Suggestions.Activate();
        }

        public void Dispose()
        {
            _suggestions?.Dispose();
        }
    }
}
