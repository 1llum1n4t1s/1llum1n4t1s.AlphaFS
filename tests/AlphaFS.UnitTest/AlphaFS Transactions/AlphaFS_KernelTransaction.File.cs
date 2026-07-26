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

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_KernelTransactionTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_WriteAllText_Commit_FileIsVisibleOutsideTransaction_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var file = tempRoot.RandomTxtFileFullPath;
            const string contents = "AlphaFS TxF commit.";

            Console.WriteLine("Input File Path: [{0}]", file);


            using (var transaction = new KernelTransaction())
            {
               File.WriteAllTextTransacted(transaction, file, contents);

               // トランザクション内からは見える。
               Assert.IsTrue(File.ExistsTransacted(transaction, file), "The file does not exist inside the transaction, but is expected to.");

               // 未コミットの変更はトランザクション外からは見えない。
               Assert.IsFalse(System.IO.File.Exists(file), "The file exists outside the uncommitted transaction, but is expected not to.");

               transaction.Commit();
            }


            Assert.IsTrue(System.IO.File.Exists(file), "The file does not exist after commit, but is expected to.");
            Assert.AreEqual(contents, System.IO.File.ReadAllText(file), "The file contents do not match, but are expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_WriteAllText_Rollback_FileIsNotCreated_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var file = tempRoot.RandomTxtFileFullPath;

            Console.WriteLine("Input File Path: [{0}]", file);


            using (var transaction = new KernelTransaction())
            {
               File.WriteAllTextTransacted(transaction, file, "AlphaFS TxF rollback.");

               Assert.IsTrue(File.ExistsTransacted(transaction, file), "The file does not exist inside the transaction, but is expected to.");

               transaction.Rollback();
            }


            Assert.IsFalse(System.IO.File.Exists(file), "The file exists after rollback, but is expected not to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_WriteAllText_DisposeWithoutCommit_FileIsNotCreated_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var file = tempRoot.RandomTxtFileFullPath;

            Console.WriteLine("Input File Path: [{0}]", file);


            // Commit も Rollback も呼ばずに Dispose した場合、TxF は自動的にロールバックする。
            using (var transaction = new KernelTransaction())
            {
               File.WriteAllTextTransacted(transaction, file, "AlphaFS TxF implicit rollback.");
            }


            Assert.IsFalse(System.IO.File.Exists(file), "The file exists after disposing the transaction without commit, but is expected not to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_ReadAllTextAndGetSize_InsideTransaction_ReflectTransactedState_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var file = tempRoot.RandomTxtFileFullPath;
            const string contents = "AlphaFS TxF read.";

            Console.WriteLine("Input File Path: [{0}]", file);


            using (var transaction = new KernelTransaction())
            {
               File.WriteAllTextTransacted(transaction, file, contents);

               Assert.AreEqual(contents, File.ReadAllTextTransacted(transaction, file), "The file contents do not match, but are expected to.");

               // UTF-8 で BOM 無し。ASCII のみなので長さはそのままバイト数になる。
               Assert.AreEqual(contents.Length, File.GetSizeTransacted(transaction, file), "The file size does not match, but is expected to.");

               transaction.Commit();
            }


            Assert.AreEqual(contents, System.IO.File.ReadAllText(file), "The file contents do not match, but are expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_Copy_Commit_DestinationIsVisibleOutsideTransaction_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var source = tempRoot.RandomTxtFileFullPath;
            var destination = tempRoot.RandomTxtFileFullPath;
            const string contents = "AlphaFS TxF copy.";

            System.IO.File.WriteAllText(source, contents);

            Console.WriteLine("Input File Path: [{0}]", source);
            Console.WriteLine("Destination File Path: [{0}]", destination);


            using (var transaction = new KernelTransaction())
            {
               // 上書きしない場合は CopyOptions.FailIfExists を使う。
               // bool overwrite や bool preserveDates を取るオーバーロードは [Obsolete] 指定されている。
               File.CopyTransacted(transaction, source, destination, CopyOptions.FailIfExists);

               Assert.IsTrue(File.ExistsTransacted(transaction, destination), "The destination file does not exist inside the transaction, but is expected to.");
               Assert.IsFalse(System.IO.File.Exists(destination), "The destination file exists outside the uncommitted transaction, but is expected not to.");

               transaction.Commit();
            }


            Assert.IsTrue(System.IO.File.Exists(source), "The source file does not exist, but is expected to.");
            Assert.IsTrue(System.IO.File.Exists(destination), "The destination file does not exist after commit, but is expected to.");
            Assert.AreEqual(contents, System.IO.File.ReadAllText(destination), "The file contents do not match, but are expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_Move_Commit_SourceIsGoneAndDestinationExists_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var source = tempRoot.RandomTxtFileFullPath;
            var destination = tempRoot.RandomTxtFileFullPath;
            const string contents = "AlphaFS TxF move.";

            System.IO.File.WriteAllText(source, contents);

            Console.WriteLine("Input File Path: [{0}]", source);
            Console.WriteLine("Destination File Path: [{0}]", destination);


            using (var transaction = new KernelTransaction())
            {
               File.MoveTransacted(transaction, source, destination);

               transaction.Commit();
            }


            Assert.IsFalse(System.IO.File.Exists(source), "The source file exists after the move, but is expected not to.");
            Assert.IsTrue(System.IO.File.Exists(destination), "The destination file does not exist after the move, but is expected to.");
            Assert.AreEqual(contents, System.IO.File.ReadAllText(destination), "The file contents do not match, but are expected to.");
         }

         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_KernelTransaction_File_Delete_Rollback_FileStillExists_Success()
      {
         using (var tempRoot = new TemporaryDirectory())
         {
            UnitTestConstants.RequireTransactionalNtfs(tempRoot.Directory.FullName);

            var file = tempRoot.RandomTxtFileFullPath;
            const string contents = "AlphaFS TxF delete rollback.";

            System.IO.File.WriteAllText(file, contents);

            Console.WriteLine("Input File Path: [{0}]", file);


            using (var transaction = new KernelTransaction())
            {
               File.DeleteTransacted(transaction, file);

               // トランザクション内では消えているが、外からはまだ見える。
               Assert.IsFalse(File.ExistsTransacted(transaction, file), "The file exists inside the transaction, but is expected not to.");
               Assert.IsTrue(System.IO.File.Exists(file), "The file does not exist outside the uncommitted transaction, but is expected to.");

               transaction.Rollback();
            }


            Assert.IsTrue(System.IO.File.Exists(file), "The file does not exist after rollback, but is expected to.");
            Assert.AreEqual(contents, System.IO.File.ReadAllText(file), "The file contents do not match, but are expected to.");
         }

         Console.WriteLine();
      }
   }
}
