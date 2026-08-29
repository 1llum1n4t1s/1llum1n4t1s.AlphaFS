using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices;
using System.Security.AccessControl;

namespace AlphaFS.UnitTest
{
   [TestClass]
   public class AlphaFS_NetworkSecurityDescriptorLifetimeTest
   {
      [TestMethod]
      public void AlphaFS_DfsInfo_CopiesSecurityDescriptorOutOfNetApiBuffer_Local_Success()
      {
         AssertDescriptorIsOwned(sourcePointer =>
         {
            var nativeInfo = new Alphaleonis.Win32.Network.NativeMethods.DFS_INFO_9
            {
               pSecurityDescriptor = sourcePointer
            };

            return new Alphaleonis.Win32.Network.DfsInfo(nativeInfo);
         });
      }


      [TestMethod]
      public void AlphaFS_ShareInfo_CopiesSecurityDescriptorOutOfNetApiBuffer_Local_Success()
      {
         AssertDescriptorIsOwned(sourcePointer =>
         {
            var nativeInfo = new Alphaleonis.Win32.Network.NativeMethods.SHARE_INFO_502
            {
               shi502_security_descriptor = sourcePointer
            };

            return new Alphaleonis.Win32.Network.ShareInfo(
               Environment.MachineName,
               Alphaleonis.Win32.Network.ShareInfoLevel.Info502,
               nativeInfo);
         });
      }


      private static void AssertDescriptorIsOwned(Func<IntPtr, object> createOwner)
      {
         var descriptor = new RawSecurityDescriptor("O:SYG:SYD:(A;;GA;;;WD)");
         var expected = new byte[descriptor.BinaryLength];
         descriptor.GetBinaryForm(expected, 0);

         var sourcePointer = Marshal.AllocHGlobal(expected.Length);
         try
         {
            Marshal.Copy(expected, 0, sourcePointer, expected.Length);
            var owner = createOwner(sourcePointer);
            var ownedPointer = owner is Alphaleonis.Win32.Network.DfsInfo dfsInfo
               ? dfsInfo.SecurityDescriptor
               : ((Alphaleonis.Win32.Network.ShareInfo) owner).SecurityDescriptor;

            Assert.AreNotEqual(IntPtr.Zero, ownedPointer);
            Assert.AreNotEqual(sourcePointer, ownedPointer);

            var actual = new byte[expected.Length];
            Marshal.Copy(ownedPointer, actual, 0, actual.Length);
            CollectionAssert.AreEqual(expected, actual);
            GC.KeepAlive(owner);
         }
         finally
         {
            Marshal.FreeHGlobal(sourcePointer);
         }
      }
   }
}
