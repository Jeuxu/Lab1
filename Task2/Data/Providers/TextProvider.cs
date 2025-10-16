using System.IO;

namespace Data.Providers
{
    public class TextProvider
    {
        public void Write(string path, string text)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Шлях до файлу не може бути порожнім.");

            File.WriteAllText(path, text);
        }
    }
}
