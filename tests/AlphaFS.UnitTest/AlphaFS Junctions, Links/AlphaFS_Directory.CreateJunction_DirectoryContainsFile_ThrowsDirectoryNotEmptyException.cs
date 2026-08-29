/*  Copyright (C) 2008-2018 Peter Palotas, Jeffrey Jangli, Alexandr Normuradov
 *  
 *  Permission is hereby granted, free of charge, to any person obtaining a copy 
 *  of this software and associated documentation files (the "Software"), to deal 
 *  in the Software without restriction, including without limitation the rights 
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
 *  copies of the Software, and to permit persons to whom the Software is 
 *  furnished to do so, subject to the following conditions:
 *  
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *  
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
 *  THE SOFTWARE. 
 */

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_JunctionsLinksTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_Directory_CreateJunction_DirectoryContainsFile_ThrowsDirectoryNotEmptyException_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var target = tempRoot.Directory.CreateSubdirectory("JunctionTarget");

         var toDelete = tempRoot.Directory.CreateSubdirectory("ToDelete");

         var junction = System.IO.Path.Combine(toDelete.FullName, "JunctionPoint");

         var dirInfo = new System.IO.DirectoryInfo(junction);
         dirInfo.Create();

         dirInfo.CreateSubdirectory("Extra Folder");
            
         UnitTestAssert.ThrowsException<Alphaleonis.Win32.Filesystem.DirectoryNotEmptyException>(() => Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junction, target.FullName));
      }


      [TestMethod]
      public void AlphaFS_Directory_CreateJunction_OverwriteOrdinaryDirectory_ThrowsNotAReparsePointException_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var target = tempRoot.Directory.CreateSubdirectory("JunctionTarget");
         var ordinaryDirectory = tempRoot.Directory.CreateSubdirectory("OrdinaryDirectory");
         var markerPath = System.IO.Path.Combine(ordinaryDirectory.FullName, "must-remain.txt");
         System.IO.File.WriteAllText(markerPath, "preserve me");

         UnitTestAssert.ThrowsException<Alphaleonis.Win32.Filesystem.NotAReparsePointException>(() =>
            Alphaleonis.Win32.Filesystem.Directory.CreateJunction(ordinaryDirectory.FullName, target.FullName, true));

         Assert.IsTrue(System.IO.Directory.Exists(ordinaryDirectory.FullName));
         Assert.AreEqual("preserve me", System.IO.File.ReadAllText(markerPath));
      }


      [TestMethod]
      public void AlphaFS_Directory_CreateJunction_OverwriteJunction_ReplacesOnlyJunction_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var firstTarget = tempRoot.Directory.CreateSubdirectory("FirstTarget");
         var secondTarget = tempRoot.Directory.CreateSubdirectory("SecondTarget");
         var junctionPath = System.IO.Path.Combine(tempRoot.Directory.FullName, "Junction");
         System.IO.File.WriteAllText(System.IO.Path.Combine(firstTarget.FullName, "first.txt"), "first");
         System.IO.File.WriteAllText(System.IO.Path.Combine(secondTarget.FullName, "second.txt"), "second");

         Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junctionPath, firstTarget.FullName);
         Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junctionPath, secondTarget.FullName, true);

         Assert.IsFalse(System.IO.File.Exists(System.IO.Path.Combine(junctionPath, "first.txt")));
         Assert.AreEqual("second", System.IO.File.ReadAllText(System.IO.Path.Combine(junctionPath, "second.txt")));
         Assert.AreEqual("first", System.IO.File.ReadAllText(System.IO.Path.Combine(firstTarget.FullName, "first.txt")));
      }


      [TestMethod]
      public void AlphaFS_Directory_CreateJunction_MissingTarget_CreatesTargetDirectory_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var targetPath = System.IO.Path.Combine(tempRoot.Directory.FullName, "MissingTarget");
         var junctionPath = System.IO.Path.Combine(tempRoot.Directory.FullName, "Junction");

         Alphaleonis.Win32.Filesystem.Directory.CreateJunction(junctionPath, targetPath);

         Assert.IsTrue(System.IO.Directory.Exists(targetPath));
         Assert.IsTrue(Alphaleonis.Win32.Filesystem.Directory.ExistsJunction(junctionPath));
      }
   }
}
