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
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] トランザクション操作として、既存のファイルと新しいファイルの間にハードリンク(CMDコマンド"MKLINK /H"と同様)を確立します。この機能はNTFSファイルシステムでのみサポートされ、ディレクトリではなくファイルのみが対象です。</summary>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="fileName">新しいファイルの名前。このパラメータではディレクトリの名前を指定できません。</param>
      /// <param name="existingFileName">既存のファイルの名前。このパラメータではディレクトリの名前を指定できません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Hardlink")]
      [SecurityCritical]
      internal static void CreateHardLinkCore(KernelTransaction transaction, string fileName, string existingFileName, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            const GetFullPathOptions options = GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck;

            fileName = Path.GetExtendedLengthPathCore(transaction, fileName, pathFormat, options);
            existingFileName = Path.GetExtendedLengthPathCore(transaction, existingFileName, pathFormat, options);
         }


         if (!(transaction == null || !NativeMethods.IsAtLeastWindowsVista

            // CreateHardLink() / CreateHardLinkTransacted()
            // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。
            // 2017-05-30: CreateHardLink() MSDNはLongPathの使用を確認: Windows 10 バージョン1607以降

            ? NativeMethods.CreateHardLink(fileName, existingFileName, IntPtr.Zero)
            : NativeMethods.CreateHardLinkTransacted(fileName, existingFileName, IntPtr.Zero, transaction.SafeHandle)))
         {
            var lastError = (uint) Marshal.GetLastWin32Error();

            switch (lastError)
            {
               case Win32Errors.ERROR_INVALID_FUNCTION:
                  throw new NotSupportedException(Resources.HardLinks_Not_Supported);

               default:
                  NativeError.ThrowException(lastError, existingFileName, fileName);
                  break;
            }
         }
      }
   }
}
