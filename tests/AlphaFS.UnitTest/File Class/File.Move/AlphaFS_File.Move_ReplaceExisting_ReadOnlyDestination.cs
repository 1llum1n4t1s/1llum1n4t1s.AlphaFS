using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   public partial class FileTest
   {
      [TestMethod]
      public void AlphaFS_File_Move_ReplaceExisting_ReadOnlyHiddenDestination_RetriesAfterClearingAttributes_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "source.txt");
         var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "destination.txt");

         System.IO.File.WriteAllText(source, "source");
         System.IO.File.WriteAllText(destination, "destination");
         System.IO.File.SetAttributes(destination, System.IO.FileAttributes.ReadOnly | System.IO.FileAttributes.Hidden);

         try
         {
            Alphaleonis.Win32.Filesystem.File.Move(
               source,
               destination,
               Alphaleonis.Win32.Filesystem.MoveOptions.ReplaceExisting);

            Assert.IsFalse(System.IO.File.Exists(source));
            Assert.AreEqual("source", System.IO.File.ReadAllText(destination));
         }
         finally
         {
            if (System.IO.File.Exists(destination))
            {
               System.IO.File.SetAttributes(destination, System.IO.FileAttributes.Normal);
            }
         }
      }
   }
}
