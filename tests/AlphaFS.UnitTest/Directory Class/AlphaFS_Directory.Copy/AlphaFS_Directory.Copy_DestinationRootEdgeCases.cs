using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   [TestClass]
   public class AlphaFS_DirectoryCopyRegressionTest
   {
      [TestMethod]
      public void AlphaFS_Directory_Copy_EmptySource_CreatesDestinationRoot_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "EmptySource");
            var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
            System.IO.Directory.CreateDirectory(source);

            var result = Alphaleonis.Win32.Filesystem.Directory.Copy(source, destination);

            Assert.IsTrue(System.IO.Directory.Exists(destination));
            Assert.AreEqual(0, result.TotalFiles);
            Assert.AreEqual(0, result.TotalFolders);
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Copy_FileOnlySource_CreatesDestinationBeforeCopyingFile_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "FileOnlySource");
            var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
            var sourceFile = System.IO.Path.Combine(source, "README.txt");
            var destinationFile = System.IO.Path.Combine(destination, "README.txt");

            System.IO.Directory.CreateDirectory(source);
            System.IO.File.WriteAllText(sourceFile, "content");

            var result = Alphaleonis.Win32.Filesystem.Directory.Copy(source, destination);

            Assert.AreEqual("content", System.IO.File.ReadAllText(destinationFile));
            Assert.AreEqual(1, result.TotalFiles);
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Copy_CopySymbolicLink_NormalDirectoryCopiesTree_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
            var destinationFile = System.IO.Path.Combine(destination, "file.txt");

            System.IO.Directory.CreateDirectory(source);
            System.IO.File.WriteAllText(System.IO.Path.Combine(source, "file.txt"), "content");

            Alphaleonis.Win32.Filesystem.Directory.Copy(
               source,
               destination,
               Alphaleonis.Win32.Filesystem.CopyOptions.CopySymbolicLink);

            Assert.AreEqual("content", System.IO.File.ReadAllText(destinationFile));
         }
      }
   }
}
