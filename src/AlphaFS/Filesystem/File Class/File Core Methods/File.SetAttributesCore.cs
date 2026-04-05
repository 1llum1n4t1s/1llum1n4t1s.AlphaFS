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
using System.IO;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>非トランザクション/トランザクションファイルまたはディレクトリの属性を設定します。</summary>
      /// <remarks>
      ///   <see cref="FileAttributes.Hidden"/>や<see cref="FileAttributes.ReadOnly"/>などの特定のファイル属性は組み合わせることができます。
      ///   <see cref="FileAttributes.Normal"/>などの他の属性は単独で使用する必要があります。
      /// </remarks>
      /// <remarks>
      ///   SetAttributesメソッドを使用してFileオブジェクトの<see cref="FileAttributes.Compressed"/>ステータスを変更することはできません。
      /// </remarks>
      /// <exception cref="ArgumentException">path is empty, contains only white spaces, contains invalid characters, or the file attribute is invalid.</exception>
      /// <exception cref="DirectoryNotFoundException">The specified path is invalid, (for example, it is on an unmapped drive).</exception>
      /// <exception cref="FileNotFoundException">The file cannot be found.</exception>
      /// <exception cref="NotSupportedException">path is in an invalid format.</exception>
      /// <exception cref="UnauthorizedAccessException">path specified a file that is read-only. -or- This operation is not supported on the current platform. -or- path specified a directory. -or- The caller does not have the required permission.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder"><paramref name="path"/>がファイルかディレクトリかを指定します。</param>
      /// <param name="path">属性を設定するファイルまたはディレクトリの名前。</param>
      /// <param name="fileAttributes">
      ///    The attributes to set for the file or directory. Note that all other values override <see cref="FileAttributes.Normal"/>.
      /// </param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static void SetAttributesCore(KernelTransaction transaction, bool isFolder, string path, FileAttributes fileAttributes, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            path = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);
         }


         var success = null == transaction || !NativeMethods.IsAtLeastWindowsVista

            // SetFileAttributes()
            // 2013-01-13: MSDNはLongPathの使用を確認しています。

            ? NativeMethods.SetFileAttributes(path, fileAttributes)

            : NativeMethods.SetFileAttributesTransacted(path, fileAttributes, transaction.SafeHandle);


         var lastError = Marshal.GetLastWin32Error();

         if (!success)
         {
            // MSDN: .NET 3.5+: ArgumentException: FileSystemInfo().Attributes

            if (lastError == Win32Errors.ERROR_INVALID_PARAMETER)

            {
               throw new ArgumentException(Resources.Invalid_File_Attribute, "fileAttributes");
            }


            NativeError.ThrowException(lastError, isFolder, path);
         }
      }
   }
}
