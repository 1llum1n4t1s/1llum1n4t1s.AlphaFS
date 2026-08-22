using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace AlphaFS.UnitTest
{
   public partial class PathTest
   {
      [TestMethod]
      public void Path_GetFullPathCore_NativePathTooLong_PropagatesWin32Error_Success()
      {
         var path = new string('a', 40000);

         var exception = Assert.ThrowsExactly<IOException>(() =>
            Alphaleonis.Win32.Filesystem.Path.GetFullPathCore(
               null,
               false,
               path,
               Alphaleonis.Win32.Filesystem.GetFullPathOptions.None));

         Assert.AreEqual(206, exception.HResult & 0xFFFF);
      }
   }
}
