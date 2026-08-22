using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.AccessControl;
using System.Security.Principal;

namespace AlphaFS.UnitTest
{
   public partial class File_CreateTest
   {
      [TestMethod]
      public void File_Create_WithDaclOnly_DoesNotRequireSecurityPrivilege_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var path = tempRoot.RandomTxtFileFullPath;
            var sid = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
            var security = new FileSecurity();
            security.AddAccessRule(new FileSystemAccessRule(
               sid,
               FileSystemRights.FullControl,
               AccessControlType.Allow));

            using (Alphaleonis.Win32.Filesystem.File.Create(
               path,
               4096,
               System.IO.FileOptions.None,
               security))
            {
            }

            var actual = Alphaleonis.Win32.Filesystem.File.GetAccessControl(path)
               .GetSecurityDescriptorSddlForm(AccessControlSections.Access);
            Assert.Contains("(A;;FA;;;WD)", actual);
         }
      }
   }
}
