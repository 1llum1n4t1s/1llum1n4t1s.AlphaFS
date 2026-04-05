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
using System.IO;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>Deletes a Non-/Transacted file.</summary>
      /// <remarks>削除するファイルが存在しない場合、例外はスローされません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <exception cref="FileReadOnlyException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">削除するファイルの名前。</param>
      /// <param name="ignoreReadOnly"><c>true</c>の場合、ファイルの読み取り専用<see cref="FileAttributes"/>を上書きします。</param>
      /// <param name="attributes">属性。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static void DeleteFileCore(KernelTransaction transaction, string path, bool ignoreReadOnly, FileAttributes attributes, PathFormat pathFormat)
      {
         if (null == path)
         {
            throw new ArgumentNullException("path");
         }

         if (pathFormat == PathFormat.RelativePath)
         {
            Path.CheckSupportedPathFormat(path, true, true);
         }

         var pathLp = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.TrimEnd | GetFullPathOptions.RemoveTrailingDirectorySeparator);


         // 既に事実がわかっている場合、属性をNormalにリセットする。

         if (ignoreReadOnly && IsReadOnlyOrHidden(attributes))

         {
            SetAttributesCore(transaction, false, pathLp, FileAttributes.Normal, PathFormat.LongFullPath);
         }


         startDeleteFile:

         if (!(null == transaction || !NativeMethods.IsAtLeastWindowsVista

            // DeleteFile() / DeleteFileTransacted()
            // 2013-01-13: MSDNはLongPathの使用を確認しています。
            //
            // パスがシンボリックリンクを指している場合、ターゲットではなくシンボリックリンクが削除されます。

            ? NativeMethods.DeleteFile(pathLp)

            : NativeMethods.DeleteFileTransacted(pathLp, transaction.SafeHandle)))
         {
            var lastError = Marshal.GetLastWin32Error();


            switch ((uint) lastError)
            {
               case Win32Errors.ERROR_FILE_NOT_FOUND:
                  // MSDN: .NET 3.5+: 削除するファイルが存在しない場合、例外はスローされません。
                  return;


               case Win32Errors.ERROR_PATH_NOT_FOUND:
                  // MSDN: .NET 3.5+: DirectoryNotFoundException: 指定されたパスが無効です(マッピングされていないドライブ上にあるなど)。
                  NativeError.ThrowException(lastError, pathLp);
                  return;


               case Win32Errors.ERROR_SHARING_VIOLATION:
                  // MSDN: .NET 3.5+: IOException: 指定されたファイルが使用中であるか、ファイルにオープンハンドルがあります。
                  NativeError.ThrowException(lastError, pathLp);
                  break;


               case Win32Errors.ERROR_ACCESS_DENIED:

                  if (attributes == 0)
                  {
                     var attrs = new NativeMethods.WIN32_FILE_ATTRIBUTE_DATA();

                     if (FillAttributeInfoCore(transaction, pathLp, ref attrs, false, true) == Win32Errors.NO_ERROR)

                     {
                        attributes = attrs.dwFileAttributes;
                     }
                  }


                  // MSDN: .NET 3.5+: UnauthorizedAccessException: Pathはディレクトリです。
                  if (IsDirectory(attributes))
                  {
                     throw new UnauthorizedAccessException(string.Format(CultureInfo.InvariantCulture, "({0}) {1}", lastError.ToString(CultureInfo.InvariantCulture), string.Format(CultureInfo.InvariantCulture, Resources.Target_File_Is_A_Directory, pathLp)));
                  }


                  if (IsReadOnlyOrHidden(attributes))
                  {
                     if (ignoreReadOnly)
                     {
                        // Reset attributes to Normal.
                        SetAttributesCore(transaction, false, pathLp, FileAttributes.Normal, PathFormat.LongFullPath);

                        goto startDeleteFile;
                     }


                     // MSDN: .NET 3.5+: UnauthorizedAccessException: Pathは読み取り専用ファイルを指定しました。
                     throw new FileReadOnlyException(pathLp);
                  }

                  
                  // MSDN: .NET 3.5+: UnauthorizedAccessException: 呼び出し元に必要なアクセス許可がありません。
                  if (attributes == 0)
                  {
                     NativeError.ThrowException(lastError, pathLp);
                  }

                  break;
            }

            // MSDN: .NET 3.5+: IOException:
            // 指定されたファイルが使用中です。
            // ファイルにオープンハンドルがあり、オペレーティングシステムがWindows XP以前です。

            NativeError.ThrowException(lastError, IsDirectory(attributes), pathLp);
         }
      }
   }
}
