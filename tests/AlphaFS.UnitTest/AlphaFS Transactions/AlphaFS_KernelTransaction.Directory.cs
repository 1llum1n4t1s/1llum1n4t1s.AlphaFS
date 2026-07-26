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

using Alphaleonis.Win32.Filesystem;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_KernelTransactionTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_KernelTransaction_Directory_CreateDirectory_Commit_DirectoryIsVisibleOutsideTransaction_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var folder = tempRoot.RandomDirectoryFullPath;

            Console.WriteLine("Input Directory Path: [{0}]", folder);


            using (var transaction = new KernelTransaction())
            {
               var dirInfo = Directory.CreateDirectoryTransacted(transaction, folder);

               Assert.IsNotNull(dirInfo, "The DirectoryInfo is null, but is expected not to be.");
               Assert.IsTrue(Directory.ExistsTransacted(transaction, folder), "The directory does not exist inside the transaction, but is expected to.");
               Assert.IsFalse(System.IO.Directory.Exists(folder), "The directory exists outside the uncommitted transaction, but is expected not to.");

               transaction.Commit();
            }


            Assert.IsTrue(System.IO.Directory.Exists(folder), "The directory does not exist after commit, but is expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_Directory_CreateDirectory_Rollback_DirectoryIsNotCreated_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var folder = tempRoot.RandomDirectoryFullPath;

            Console.WriteLine("Input Directory Path: [{0}]", folder);


            using (var transaction = new KernelTransaction())
            {
               Directory.CreateDirectoryTransacted(transaction, folder);

               Assert.IsTrue(Directory.ExistsTransacted(transaction, folder), "The directory does not exist inside the transaction, but is expected to.");

               transaction.Rollback();
            }


            Assert.IsFalse(System.IO.Directory.Exists(folder), "The directory exists after rollback, but is expected not to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_Directory_CreateDirectory_NestedPath_Commit_AllLevelsExist_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var level1 = tempRoot.RandomDirectoryFullPath;
            var level3 = System.IO.Path.Combine(level1, tempRoot.RandomDirectoryName, tempRoot.RandomDirectoryName);

            Console.WriteLine("Input Directory Path: [{0}]", level3);


            using (var transaction = new KernelTransaction())
            {
               // 途中のディレクトリがすべて存在しない状態から作る。
               // 通常版の CreateDirectoryCore と同じ祖先走査を通るため、そこへの回帰も検知できる。
               Directory.CreateDirectoryTransacted(transaction, level3);

               transaction.Commit();
            }


            Assert.IsTrue(System.IO.Directory.Exists(level1), "The first level directory does not exist, but is expected to.");
            Assert.IsTrue(System.IO.Directory.Exists(level3), "The third level directory does not exist, but is expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_Directory_EnumerateFiles_InsideTransaction_SeesTransactedFiles_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var folder = tempRoot.Directory.FullName;
            var committedFile = tempRoot.RandomTxtFileFullPath;
            var transactedFile = tempRoot.RandomTxtFileFullPath;

            System.IO.File.WriteAllText(committedFile, "committed");

            Console.WriteLine("Input Directory Path: [{0}]", folder);


            using (var transaction = new KernelTransaction())
            {
               File.WriteAllTextTransacted(transaction, transactedFile, "transacted");

               var transactedView = Directory.EnumerateFilesTransacted(transaction, folder).ToArray();

               CollectionAssert.Contains(transactedView, committedFile, "The committed file is not enumerated, but is expected to be.");
               CollectionAssert.Contains(transactedView, transactedFile, "The transacted file is not enumerated, but is expected to be.");

               // トランザクション外の列挙にはまだ現れない。
               var outsideView = System.IO.Directory.GetFiles(folder);

               CollectionAssert.Contains(outsideView, committedFile, "The committed file is not enumerated outside the transaction, but is expected to be.");
               CollectionAssert.DoesNotContain(outsideView, transactedFile, "The transacted file is enumerated outside the uncommitted transaction, but is expected not to be.");

               transaction.Commit();
            }


            Assert.IsTrue(System.IO.File.Exists(transactedFile), "The transacted file does not exist after commit, but is expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_Directory_Delete_Rollback_DirectoryStillExists_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var folder = tempRoot.RandomDirectoryFullPath;

            System.IO.Directory.CreateDirectory(folder);

            Console.WriteLine("Input Directory Path: [{0}]", folder);


            using (var transaction = new KernelTransaction())
            {
               Directory.DeleteTransacted(transaction, folder);

               Assert.IsFalse(Directory.ExistsTransacted(transaction, folder), "The directory exists inside the transaction, but is expected not to.");
               Assert.IsTrue(System.IO.Directory.Exists(folder), "The directory does not exist outside the uncommitted transaction, but is expected to.");

               transaction.Rollback();
            }


            Assert.IsTrue(System.IO.Directory.Exists(folder), "The directory does not exist after rollback, but is expected to.");
         }

         Console.WriteLine();
      }
   }
}
