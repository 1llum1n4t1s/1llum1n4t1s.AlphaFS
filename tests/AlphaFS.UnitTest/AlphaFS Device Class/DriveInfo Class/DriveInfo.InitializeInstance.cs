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

using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   public partial class DriveInfoTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void DriveInfo_InitializeInstance_LocalAndNetwork_Success()
      {
         DriveInfo_InitializeInstance(false);
         DriveInfo_InitializeInstance(true);
      }
      

      private void DriveInfo_InitializeInstance(bool isNetwork)
      {
         UnitTestConstants.PrintUnitTestHeader(isNetwork);
         
         var drive = UnitTestConstants.SysDrive[0].ToString();

         if (isNetwork)
            // Only using a drive letter results in a wrong UNC path.
         {
            drive = Alphaleonis.Win32.Filesystem.Path.LocalToUnc(UnitTestConstants.SysDrive);
         }

         Console.WriteLine("Input Drive Path: [{0}]", drive);


         var actual = new Alphaleonis.Win32.Filesystem.DriveInfo(drive);

         Assert.IsTrue(actual.IsReady);
         Assert.IsTrue(actual.IsVolume);

         if (isNetwork)
         {
            Assert.IsTrue(actual.IsUnc);
         }
         else
         {
            Assert.IsFalse(actual.IsUnc);
         }


         // System.IO.DriveInfo cannot handle UNC paths.

         if (!isNetwork)
         {
            var expected = new System.IO.DriveInfo(drive);


            // 空き容量は 2 回の読み取りの間にも変動する (テストは並列実行され、他プロセスも書き込む)。
            // 完全一致を要求すると本質的に不安定になるため、実装差を検知できる範囲の許容誤差で比較する。
            AssertFreeSpaceClose(expected.AvailableFreeSpace, actual.AvailableFreeSpace, expected.TotalSize, "AvailableFreeSpace");
            AssertFreeSpaceClose(expected.TotalFreeSpace, actual.TotalFreeSpace, expected.TotalSize, "TotalFreeSpace");

            // 総容量は変動しないので完全一致を要求する。
            Assert.AreEqual(expected.TotalSize, actual.TotalSize, "TotalSize AlphaFS != System.IO");


            Assert.AreEqual(expected.DriveFormat, actual.DriveFormat, "DriveFormat AlphaFS != System.IO");
            Assert.AreEqual(expected.DriveType, actual.DriveType, "DriveType AlphaFS != System.IO");
            Assert.AreEqual(expected.IsReady, actual.IsReady, "IsReady AlphaFS != System.IO");
            Assert.AreEqual(expected.Name, actual.Name, "Name AlphaFS != System.IO");
            Assert.AreEqual(expected.RootDirectory.ToString(), actual.RootDirectory.ToString(), "RootDirectory AlphaFS != System.IO");
            Assert.AreEqual(expected.VolumeLabel, actual.VolumeLabel, "VolumeLabel AlphaFS != System.IO");


            UnitTestConstants.Dump(expected);
            Console.WriteLine();
         }


         UnitTestConstants.Dump(actual);

         UnitTestConstants.Dump(actual.DiskSpaceInfo);

         UnitTestConstants.Dump(actual.VolumeInfo);

         Console.WriteLine();
      }


      /// <summary>空き容量を許容誤差付きで比較する。実装差 (0 や別ボリュームの値) は検知しつつ、読み取り間の変動は許容する。</summary>
      private static void AssertFreeSpaceClose(long expected, long actual, long totalSize, string propertyName)
      {
         // ドライブ全体の 1% か 256 MB の大きい方を許容幅とする。
         var tolerance = Math.Max(totalSize / 100, 256L * 1024 * 1024);
         var difference = Math.Abs(expected - actual);

         Console.WriteLine("\t{0}: System.IO=[{1:N0}] AlphaFS=[{2:N0}] 差=[{3:N0}] 許容=[{4:N0}]", propertyName, expected, actual, difference, tolerance);

         Assert.IsLessThanOrEqualTo(tolerance, difference,
            string.Format(CultureInfo.CurrentCulture, "{0} の差が許容範囲を超えています。AlphaFS != System.IO", propertyName));
      }
   }
}
