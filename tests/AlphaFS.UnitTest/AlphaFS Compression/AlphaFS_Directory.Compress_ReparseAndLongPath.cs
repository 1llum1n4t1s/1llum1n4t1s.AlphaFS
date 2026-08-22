using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace AlphaFS.UnitTest
{
   partial class AlphaFS_CompressionTest
   {
      [TestMethod]
      public void AlphaFS_Directory_Compress_Recursive_DoesNotFollowJunction_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var targetFile = System.IO.Path.Combine(target, "target.txt");
            var junction = System.IO.Path.Combine(source, "TargetJunction");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(targetFile, "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junction, target);

            try
            {
               Alphaleonis.Win32.Filesystem.Directory.Compress(
                  source,
                  Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.Recursive);

               FileAssert.IsNotCompressed(targetFile);
            }
            finally
            {
               if (System.IO.File.Exists(targetFile) &&
                   (System.IO.File.GetAttributes(targetFile) & FileAttributes.Compressed) != 0)
               {
                  Alphaleonis.Win32.Filesystem.File.Decompress(targetFile);
               }
            }
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Compress_Recursive_UsesLongFullPathForChildren_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var root = System.IO.Path.Combine(tempRoot.Directory.FullName, "LongPathRoot");
            var deepest = root;

            while (deepest.Length <= 300)
            {
               deepest = System.IO.Path.Combine(deepest, "123456789012345678901234567890");
            }

            Alphaleonis.Win32.Filesystem.Directory.CreateDirectory(deepest);
            var file = System.IO.Path.Combine(deepest, "file.txt");
            System.IO.File.WriteAllText(file, "content");

            Alphaleonis.Win32.Filesystem.Directory.Compress(
               root,
               Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.Recursive);

            FileAssert.IsCompressed(file);
         }
      }
   }
}
