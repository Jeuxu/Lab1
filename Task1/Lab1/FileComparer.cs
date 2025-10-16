using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Lab1
{
    public static class FileComparer
    {
        public static bool AreFilesEqual(string path1, string path2)
        {
            byte[] hash1 = GetFileHash(path1);
            byte[] hash2 = GetFileHash(path2);

            return StructuralComparisons.StructuralEqualityComparer.Equals(hash1, hash2);
        }

        private static byte[] GetFileHash(string path)
        {
            using (var md5 = MD5.Create())
            using (var stream = File.OpenRead(path))
            {
                return md5.ComputeHash(stream);
            }
        }
    }
}
