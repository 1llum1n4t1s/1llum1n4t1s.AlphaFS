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
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>ファイルまたはディレクトリが作成された日時、および/または最終アクセス日時、および/または最終書き込み日時を協定世界時(UTC)で設定します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder"><paramref name="path"/>がファイルかディレクトリかを指定します。</param>
      /// <param name="path">日時情報を設定するファイルまたはディレクトリ。</param>
      /// <param name="creationTimeUtc"><paramref name="path"/>の作成日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastAccessTimeUtc"><paramref name="path"/>の最終アクセス日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastWriteTimeUtc"><paramref name="path"/>の最終書き込み日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="modifyReparsePoint"><c>true</c>の場合、日時情報はリパースポイント(シンボリックリンクまたはジャンクション)に適用され、リンク先のファイルまたはディレクトリには適用されません。<paramref name="path"/>がリパースポイントを参照していない場合は効果がありません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static void SetFsoDateTimeCore(KernelTransaction transaction, bool isFolder, string path, DateTime? creationTimeUtc, DateTime? lastAccessTimeUtc, DateTime? lastWriteTimeUtc, bool modifyReparsePoint, PathFormat pathFormat)
      {
         if (pathFormat == PathFormat.RelativePath)
         {
            Path.CheckSupportedPathFormat(path, false, false);
         }

         var eaAttributes = ExtendedFileAttributes.BackupSemantics;

         if (modifyReparsePoint)
         {
            eaAttributes |= ExtendedFileAttributes.OpenReparsePoint;
         }


         //// Prevent a System.UnauthorizedAccessException from being thrown by resetting attributes to Normal.

         //var fileAttributes = FileAttributes.Normal;
         //var isReadOnly = IsReadOnly((FileAttributes) eaAttributes);
         //var isHidden = IsHidden((FileAttributes) eaAttributes);

         //if (isReadOnly || isHidden)
         //   SetAttributesCore(transaction, isFolder, path, fileAttributes, PathFormat.LongFullPath);


         using var creationTime = SafeGlobalMemoryBufferHandle.FromLong(creationTimeUtc.HasValue ? creationTimeUtc.Value.ToFileTimeUtc() : (long?) null);
         using var lastAccessTime = SafeGlobalMemoryBufferHandle.FromLong(lastAccessTimeUtc.HasValue ? lastAccessTimeUtc.Value.ToFileTimeUtc() : (long?) null);
         using var lastWriteTime = SafeGlobalMemoryBufferHandle.FromLong(lastWriteTimeUtc.HasValue ? lastWriteTimeUtc.Value.ToFileTimeUtc() : (long?) null);
         using var safeFileHandle = CreateFileCore(transaction, isFolder, path, eaAttributes, null, FileMode.Open, FileSystemRights.WriteAttributes, FileShare.Delete | FileShare.Write, false, false, pathFormat);
         var success = NativeMethods.SetFileTime(safeFileHandle, creationTime, lastAccessTime, lastWriteTime);

         var lastError = Marshal.GetLastWin32Error();
            
         // ファイルシステムオブジェクトの属性をリセットする。

         if (success)
         {
            //if (isReadOnly || isHidden)
            //{
            //   if (isReadOnly)
            //      fileAttributes |= FileAttributes.ReadOnly;

            //   if (isHidden)
            //      fileAttributes |= FileAttributes.Hidden;

            //   fileAttributes &= ~FileAttributes.Normal;

            //   SetAttributesCore(transaction, isFolder, path, fileAttributes, PathFormat.LongFullPath);
            //}
         }

         else
         {
            NativeError.ThrowException(lastError, isFolder, path);
         }
      }
   }
}
