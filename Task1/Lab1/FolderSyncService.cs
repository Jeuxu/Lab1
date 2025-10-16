using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab1
{
    public class FolderSyncService
    {
        public event Action<string> LogGenerated;

        public void Sync(string source, string target)
        {
            Directory.CreateDirectory(target);

            // копіювання та оновлення
            foreach (string file in Directory.GetFiles(source))
            {
                string fileName = Path.GetFileName(file);
                string targetFile = Path.Combine(target, fileName);

                if (!File.Exists(targetFile) || !FileComparer.AreFilesEqual(file, targetFile))
                {
                    File.Copy(file, targetFile, true);
                    LogGenerated?.Invoke($"Скопійовано/оновлено: {fileName}");
                }
            }

            // видалення зайвих
            foreach (string file in Directory.GetFiles(target))
            {
                string fileName = Path.GetFileName(file);
                string sourceFile = Path.Combine(source, fileName);

                if (!File.Exists(sourceFile))
                {
                    File.Delete(file);
                    LogGenerated?.Invoke($"Видалено: {fileName}");
                }
            }
        }
    }
}
