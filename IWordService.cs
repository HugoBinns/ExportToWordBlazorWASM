namespace ExportToWord.Services.IServices
{
    public interface IWordService
    {
        byte[] GenerateWordDocument(byte[] templateBytes, Dictionary<string, string> parameters);
    }

}
