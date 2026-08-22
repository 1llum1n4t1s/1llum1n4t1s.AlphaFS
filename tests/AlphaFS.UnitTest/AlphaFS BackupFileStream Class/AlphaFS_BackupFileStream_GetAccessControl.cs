using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Security.AccessControl;

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_BackupFileStreamTest
   {
      [TestMethod]
      public void AlphaFS_BackupFileStream_GetAccessControl_PInvokeDescriptor_Local_Success()
      {
         UnitTestAssert.IsElevatedProcess();

         using (var tempRoot = new TemporaryDirectory())
         {
            var file = tempRoot.CreateFile();

            using (new Alphaleonis.Win32.Security.PrivilegeEnabler(Alphaleonis.Win32.Security.Privilege.Security))
            using (var stream = new Alphaleonis.Win32.Filesystem.BackupFileStream(
               file.FullName,
               FileMode.Open,
               FileSystemRights.Read | (FileSystemRights) 0x01000000,
               FileShare.ReadWrite))
            {
               var security = stream.GetAccessControl();

               Assert.IsNotNull(security);
               Assert.IsFalse(string.IsNullOrEmpty(
                  security.GetSecurityDescriptorSddlForm(AccessControlSections.Access)));
            }
         }
      }
   }
}
