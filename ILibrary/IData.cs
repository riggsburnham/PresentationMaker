namespace ILibrary
{
    // using this interface purely for extensibility, in the future can add more apis to call from.
    // all they would need to do is implement this interface similar to GoogleLibrary
    public interface IData
    {
        string URL { get; set; }
    }
}
