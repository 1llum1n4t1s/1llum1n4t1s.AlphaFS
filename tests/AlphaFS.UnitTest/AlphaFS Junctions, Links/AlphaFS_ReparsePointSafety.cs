using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace AlphaFS.UnitTest
{
   [TestClass]
   public class AlphaFS_ReparsePointSafetyTest
   {
      [TestMethod]
      public void AlphaFS_Directory_Delete_Recursive_DoesNotFollowDirectorySymbolicLink_Local_Success()
      {
         UnitTestAssert.IsElevatedProcess();

         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var targetFile = System.IO.Path.Combine(target, "target.txt");
            var link = System.IO.Path.Combine(source, "TargetLink");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(targetFile, "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateSymbolicLink(link, target);

            Alphaleonis.Win32.Filesystem.Directory.Delete(source, true);

            Assert.IsFalse(System.IO.Directory.Exists(source));
            Assert.IsTrue(System.IO.Directory.Exists(target));
            Assert.IsTrue(System.IO.File.Exists(targetFile));
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_DeleteEmptySubdirectories_Recursive_DoesNotFollowDirectorySymbolicLink_Local_Success()
      {
         UnitTestAssert.IsElevatedProcess();

         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var targetChild = System.IO.Path.Combine(target, "EmptyChild");
            var link = System.IO.Path.Combine(source, "TargetLink");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(targetChild);
            Alphaleonis.Win32.Filesystem.Directory.CreateSymbolicLink(link, target);

            Alphaleonis.Win32.Filesystem.Directory.DeleteEmptySubdirectories(source, true);

            Assert.IsTrue(System.IO.Directory.Exists(targetChild));
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_DeleteEmptySubdirectories_Recursive_DoesNotFollowJunction_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var targetChild = System.IO.Path.Combine(target, "EmptyChild");
            var junction = System.IO.Path.Combine(source, "TargetJunction");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(targetChild);
            Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junction, target);

            Alphaleonis.Win32.Filesystem.Directory.DeleteEmptySubdirectories(source, true);

            Assert.IsTrue(System.IO.Directory.Exists(targetChild));
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Copy_CopySymbolicLink_PreservesNestedDirectorySymbolicLink_Local_Success()
      {
         UnitTestAssert.IsElevatedProcess();

         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var sourceLink = System.IO.Path.Combine(source, "TargetLink");
            var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
            var destinationLink = System.IO.Path.Combine(destination, "TargetLink");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(System.IO.Path.Combine(target, "target.txt"), "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateSymbolicLink(sourceLink, target);

            Alphaleonis.Win32.Filesystem.Directory.Copy(
               source,
               destination,
               Alphaleonis.Win32.Filesystem.CopyOptions.CopySymbolicLink);

            var sourceInfo = Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(sourceLink);
            var destinationInfo = Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(destinationLink);

            Assert.IsTrue(new Alphaleonis.Win32.Filesystem.DirectoryInfo(destinationLink).EntryInfo.IsSymbolicLink);
            Assert.AreEqual(sourceInfo.SubstituteName, destinationInfo.SubstituteName);
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Copy_NestedJunction_PreservesJunctionWithoutMaterializingTarget_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var sourceJunction = System.IO.Path.Combine(source, "TargetJunction");
            var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
            var destinationJunction = System.IO.Path.Combine(destination, "TargetJunction");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(System.IO.Path.Combine(target, "target.txt"), "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateJunction(sourceJunction, target);

            Alphaleonis.Win32.Filesystem.Directory.Copy(source, destination);

            var destinationInfo = new Alphaleonis.Win32.Filesystem.DirectoryInfo(destinationJunction).EntryInfo;
            Assert.IsTrue(destinationInfo.IsMountPoint);
            Assert.AreEqual(
               Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(sourceJunction).PrintName,
               Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(destinationJunction).PrintName,
               true);
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Copy_RootJunction_PreservesJunctionWithoutMaterializingTarget_Local_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var sourceJunction = System.IO.Path.Combine(tempRoot.Directory.FullName, "SourceJunction");
            var destinationJunction = System.IO.Path.Combine(tempRoot.Directory.FullName, "DestinationJunction");

            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(System.IO.Path.Combine(target, "target.txt"), "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateJunction(sourceJunction, target);

            Alphaleonis.Win32.Filesystem.Directory.Copy(sourceJunction, destinationJunction);

            Assert.IsTrue(new Alphaleonis.Win32.Filesystem.DirectoryInfo(destinationJunction).EntryInfo.IsMountPoint);
            Assert.AreEqual(
               Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(sourceJunction).PrintName,
               Alphaleonis.Win32.Filesystem.Directory.GetLinkTargetInfo(destinationJunction).PrintName,
               true);
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Encrypt_Recursive_DoesNotFollowDirectorySymbolicLink_Local_Success()
      {
         UnitTestAssert.IsElevatedProcess();

         using (var tempRoot = new TemporaryDirectory())
         {
            var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
            var target = System.IO.Path.Combine(tempRoot.Directory.FullName, "Target");
            var targetFile = System.IO.Path.Combine(target, "target.txt");
            var link = System.IO.Path.Combine(source, "TargetLink");

            System.IO.Directory.CreateDirectory(source);
            System.IO.Directory.CreateDirectory(target);
            System.IO.File.WriteAllText(targetFile, "target");
            Alphaleonis.Win32.Filesystem.Directory.CreateSymbolicLink(link, target);

            try
            {
               Alphaleonis.Win32.Filesystem.Directory.Encrypt(source, true);

               Assert.AreEqual(
                  (FileAttributes) 0,
                  System.IO.File.GetAttributes(targetFile) & FileAttributes.Encrypted,
                  "リンク先のファイルは暗号化されてはいけません。");
            }
            finally
            {
               if (System.IO.File.Exists(targetFile) &&
                   (System.IO.File.GetAttributes(targetFile) & FileAttributes.Encrypted) != 0)
               {
                  Alphaleonis.Win32.Filesystem.File.Decrypt(targetFile);
               }
            }
         }
      }


      [TestMethod]
      public void AlphaFS_Directory_Encrypt_Recursive_DoesNotFollowJunction_Local_Success()
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
               Alphaleonis.Win32.Filesystem.Directory.Encrypt(source, true);

               Assert.AreEqual(
                  (FileAttributes) 0,
                  System.IO.File.GetAttributes(targetFile) & FileAttributes.Encrypted,
                  "ジャンクション先のファイルは暗号化されてはいけません。");
            }
            finally
            {
               if (System.IO.File.Exists(targetFile) &&
                   (System.IO.File.GetAttributes(targetFile) & FileAttributes.Encrypted) != 0)
               {
                  Alphaleonis.Win32.Filesystem.File.Decrypt(targetFile);
               }
            }
         }
      }
   }
}
