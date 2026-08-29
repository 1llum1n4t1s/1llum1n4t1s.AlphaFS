using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace AlphaFS.UnitTest
{
   [TestClass]
   public class AlphaFS_FileInfoCopyLongFullPathOptionsTest
   {
      [TestMethod]
      public void AlphaFS_FileInfo_Copy_PreserveDatesAndProgress_LongFullPathPipeline_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "source.bin");
         var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "destination.bin");
         var expectedLastWriteTime = new DateTime(2020, 5, 6, 7, 8, 10, DateTimeKind.Utc);
         var callbackCount = 0;

         System.IO.File.WriteAllBytes(source, new byte[128 * 1024]);
         System.IO.File.SetLastWriteTimeUtc(source, expectedLastWriteTime);

         var sourceInfo = new Alphaleonis.Win32.Filesystem.FileInfo(source);
         sourceInfo.CopyTo(
            destination,
            Alphaleonis.Win32.Filesystem.CopyOptions.None,
            true,
            (totalFileSize, totalBytesTransferred, streamSize, streamBytesTransferred, streamNumber, callbackReason, userData) =>
            {
               callbackCount++;
               return Alphaleonis.Win32.Filesystem.CopyMoveProgressResult.Continue;
            },
            null);

         Assert.IsGreaterThan(0, callbackCount);
         Assert.AreEqual(expectedLastWriteTime, System.IO.File.GetLastWriteTimeUtc(destination));
      }
   }
}
