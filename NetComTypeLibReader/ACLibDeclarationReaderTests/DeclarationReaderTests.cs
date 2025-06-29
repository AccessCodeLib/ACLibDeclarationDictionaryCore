using System;
using System.IO;
using Xunit;
using AccessCodeLib.DeclarationReader;

namespace AccessCodeLib.DeclarationReader.TypeLibReaderTests
{
    public class DeclarationReaderTests
    {
        [Fact]
        public void ReadDaoLib()
        {
            // Typical DAO library path for 64-bit Office installations
            // Adjust the path if using 32-bit Office or a different version
            string daoLibPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                @"Common Files\Microsoft Shared\DAO\dao360.dll");

            if (!File.Exists(daoLibPath))
            {
                // Skip the test if the DAO library is not found
                throw new FileNotFoundException($"DAO library not found at: {daoLibPath}");
            }

            var reader = new TypeLibDeclarationReader();
            var declarations = reader.ReadTypeLib(daoLibPath);

            Assert.NotNull(declarations);
            Assert.NotEmpty(declarations);

            // Expected words to check in the declarations
            string[] expectedWords = { "DAO", "Database", "IdleEnum", "dbRefreshCache", "Field", "OrdinalPosition", "Refresh" };

            foreach (var word in expectedWords)
            {
                Assert.Contains(declarations, d => d.Contains(word, StringComparison.OrdinalIgnoreCase));
            }
        }
    }
}