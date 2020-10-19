using System;

namespace ILibrary
{
    public interface IData
    {
        string URL { get; set; }

        // TODO: could create a list of strings that would hold extra data to be displayed...
        // IList<string> ImageData();
    }
}
