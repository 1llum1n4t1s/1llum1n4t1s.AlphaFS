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
using System;

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_VolumeTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_Volume_GetVolumeLabel_Local_Success()
      {
         UnitTestConstants.PrintUnitTestHeader(false);


         var logicalDriveCount = 0;

         foreach (var driveInfo in System.IO.DriveInfo.GetDrives())
         {
            if (!driveInfo.IsReady)
            {
               continue;
            }


            // System.IO 側のラベルを先に 1 回だけ読む。
            //
            // GetDrives() での列挙、IsReady の確認、VolumeLabel の読み取りは別々の時点なので、
            // その間にドライブが消えると DriveNotFoundException になる。
            // ネットワークドライブの切断のほか、ドライブ文字を割り当てる他のテスト
            // (AlphaFS_Host.ConnectDrive_And_DisconnectDrive は「最後の空き文字」= 通常 Z: を使う)
            // と重なった場合にも起こる。
            //
            // 比較対象の期待値を取れないだけなので、そのドライブは対象から外す。
            // ガードは System.IO 側の probe だけに掛け、被テスト対象である
            // Volume.GetVolumeLabel の失敗は握り潰さない。
            string expectedLabel;

            try
            {
               expectedLabel = driveInfo.VolumeLabel;
            }
            catch (System.IO.IOException ex)
            {
               Console.WriteLine("#{0:000}\tSkipped Logical Drive Path: [{1}]\t\t{2}", logicalDriveCount + 1, driveInfo.Name, ex.Message);

               continue;
            }


            Console.Write("#{0:000}\tInput Logical Drive Path: [{1}]", ++logicalDriveCount, driveInfo.Name);


            string volumeLabel;

            try
            {
               volumeLabel = Alphaleonis.Win32.Filesystem.Volume.GetVolumeLabel(driveInfo.Name);
            }
            catch (System.IO.IOException)
            {
               // AlphaFS 側が失敗した場合、ドライブがこの瞬間に消えたのかを System.IO で再確認する。
               // System.IO でも読めなくなっていれば環境要因なので skip、読めるなら本物の回帰として投げ直す。
               if (!DriveHasVanished(driveInfo))
               {
                  throw;
               }

               Console.WriteLine("\t\t(drive became unavailable, skipped)");

               --logicalDriveCount;

               continue;
            }


            Console.WriteLine("\t\tLabel: [{0}]", expectedLabel);


            Assert.AreEqual(expectedLabel, volumeLabel, "The volume labels do not match, but it is expected.");
         }


         Assert.IsGreaterThan(0, logicalDriveCount, "No logical drives enumerated, but it is expected.");
      }


      /// <summary>ドライブが列挙後に利用できなくなったかを System.IO で再確認します。</summary>
      /// <returns>ドライブが消えている場合は <c>true</c>。まだ読める場合は <c>false</c>。</returns>
      private static bool DriveHasVanished(System.IO.DriveInfo driveInfo)
      {
         try
         {
            var recheck = new System.IO.DriveInfo(driveInfo.Name);

            if (!recheck.IsReady)
            {
               return true;
            }

            var unused = recheck.VolumeLabel;

            return false;
         }
         catch (System.IO.IOException)
         {
            return true;
         }
      }
   }
}
