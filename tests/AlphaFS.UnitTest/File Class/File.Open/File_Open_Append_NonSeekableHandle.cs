using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32.SafeHandles;
using System.IO;
using System.Threading;

namespace AlphaFS.UnitTest
{
   public partial class FileTest
   {
      [TestMethod]
      public void AlphaFS_File_SetFilePointerForAppend_NonSeekableHandle_ClosesHandleAndThrowsIOException_Local_Success()
      {
         using var waitHandle = new EventWaitHandle(false, EventResetMode.AutoReset);
         using var handle = new SafeFileHandle(waitHandle.SafeWaitHandle.DangerousGetHandle(), false);

         Assert.ThrowsExactly<IOException>(() =>
            Alphaleonis.Win32.Filesystem.File.SetFilePointerForAppend(
               handle,
               false,
               "event handle",
               false));

         Assert.IsTrue(handle.IsClosed);
      }
   }
}
