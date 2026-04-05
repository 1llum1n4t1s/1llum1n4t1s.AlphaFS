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
      /// <summary>ファイルまたはディレクトリを復号化/暗号化し、暗号化に使用したアカウントのみが復号化できるようにします。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryReadOnlyException"/>
      /// <exception cref="FileReadOnlyException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="isFolder"><paramref name="path"/>がファイルかディレクトリかを指定します。</param>
      /// <param name="path">暗号化するファイルを示すパス。</param>
      /// <param name="encrypt"><c>true</c>で暗号化、<c>false</c>で復号化。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static void EncryptDecryptFileCore(bool isFolder, string path, bool encrypt, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            path = Path.GetExtendedLengthPathCore(null, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

            pathFormat = PathFormat.LongFullPath;
         }


         // MSDN: lpFileNameが読み取り専用ファイルを指定した場合、関数は失敗し、GetLastErrorはERROR_FILE_READ_ONLYを返します。

         var attrs = GetAttributesExCore<NativeMethods.WIN32_FILE_ATTRIBUTE_DATA>(null, path, pathFormat, true);

         var isReadOnly = IsReadOnly(attrs.dwFileAttributes);
         var isHidden = IsHidden(attrs.dwFileAttributes);

         if (isReadOnly || isHidden)
         {
            if (isReadOnly)
            {
               attrs.dwFileAttributes &= ~FileAttributes.ReadOnly;
            }

            if (isHidden)
            {
               attrs.dwFileAttributes &= ~FileAttributes.Hidden;
            }

            SetAttributesCore(null, isFolder, path, attrs.dwFileAttributes, pathFormat);
         }


         // EncryptFile() / DecryptFile()
         // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。

         var success = encrypt ? NativeMethods.EncryptFile(path) : NativeMethods.DecryptFile(path, 0);

         var lastError = Marshal.GetLastWin32Error();


         if (isReadOnly || isHidden)
         {
            if (isReadOnly)
            {
               attrs.dwFileAttributes |= FileAttributes.ReadOnly;
            }

            if (isHidden)
            {
               attrs.dwFileAttributes |= FileAttributes.Hidden;
            }

            SetAttributesCore(null, isFolder, path, attrs.dwFileAttributes, pathFormat);
         }


         if (!success)
         {
            switch ((uint) lastError)
            {
               case Win32Errors.ERROR_ACCESS_DENIED:

                  if (!string.Equals("NTFS", new DriveInfo(path).DriveFormat, StringComparison.OrdinalIgnoreCase))

                  {
                     throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "The drive does not support NTFS encryption: [{0}]", Path.GetPathRoot(path, false)));
                  }

                  break;


               case Win32Errors.ERROR_FILE_READ_ONLY:

                  if (isFolder)
                  {
                     throw new DirectoryReadOnlyException(path);
                  }

                  else
                  {
                     throw new FileReadOnlyException(path);
                  }


               default:
                  NativeError.ThrowException(lastError, isFolder, path);
                  break;
            }
         }
      }
   }
}
