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

using Alphaleonis.Win32.Security;
using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>Creates or opens a file, directory or I/O device.</summary>
      /// <returns><paramref name="path"/>で指定されたファイルまたはディレクトリへの読み取り/書き込みアクセスを提供する<see cref="SafeFileHandle"/>。</returns>
      /// <remarks>
      ///   <para>To obtain a directory handle using CreateFile, specify the FILE_FLAG_BACKUP_SEMANTICS flag as part of dwFlagsAndAttributes.</para>
      ///   <para>The most commonly used I/O devices are as follows: file, file stream, directory, physical disk, volume, console buffer, tape drive, communications resource, mailslot, and pipe.</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="Exception"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="isFolder">When <c>true</c> indicates the source is a directory, <c>false</c> indicates a file and <c>null</c> specifies a physical device.</param>
      /// <param name="path">The path and name of the file or directory to create.</param>
      /// <param name="attributes">One of the <see cref="ExtendedFileAttributes"/> values that describes how to create or overwrite the file or directory.</param>
      /// <param name="fileSecurity">ファイルまたはディレクトリのアクセス制御と監査セキュリティを決定する<see cref="FileSecurity"/>インスタンス。</param>
      /// <param name="fileMode">ファイルまたはディレクトリの開き方または作成方法を決定する<see cref="FileMode"/>定数。</param>
      /// <param name="fileSystemRights">ファイルまたはディレクトリのアクセスルールと監査ルールの作成時に使用するアクセス権を決定する<see cref="FileSystemRights"/>定数。</param>
      /// <param name="fileShare">プロセスによるファイルまたはディレクトリの共有方法を決定する<see cref="FileShare"/>定数。</param>
      /// <param name="checkPath">チェックするパス。</param>
      /// <param name="continueOnException"><c>true</c>の場合、ACLで保護されたディレクトリやアクセスできないリパースポイントなどの失敗の結果としてスローされる可能性のある例外を抑制します。</param>
      /// <param name="pathFormat"><paramref name="path"/>パラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Object needs to be disposed by caller.")]
      [SecurityCritical]
      internal static SafeFileHandle CreateFileCore(KernelTransaction transaction, bool? isFolder, string path, ExtendedFileAttributes attributes, FileSecurity fileSecurity, FileMode fileMode, FileSystemRights fileSystemRights, FileShare fileShare, bool checkPath, bool continueOnException, PathFormat pathFormat)
      {
         if (checkPath && pathFormat == PathFormat.RelativePath)

         {
            Path.CheckSupportedPathFormat(path, true, true);
         }


         // isFile == nullの場合、デバイスを操作しています。
         // ボリュームまたはリムーバブルメディアドライブ(フロッピーディスクドライブやフラッシュメモリサムドライブなど)を開く場合、
         // パス文字列は次の形式にする必要があります: "\\.\X:"
         // ルートを示す末尾のバックスラッシュ('\')は使用しないでください。

         var pathLp = Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.TrimEnd | GetFullPathOptions.RemoveTrailingDirectorySeparator);


         // CreateFileXxx()はFileMode.Appendモードをサポートしていません。
         var isAppend = fileMode == FileMode.Append;
         if (isAppend)
         {
            fileMode = FileMode.OpenOrCreate;
            fileSystemRights |= FileSystemRights.AppendData;
         }


         if (null != fileSecurity)
         {
            fileSystemRights |= (FileSystemRights) SECURITY_INFORMATION.UNPROTECTED_SACL_SECURITY_INFORMATION;
         }


         using ((fileSystemRights & (FileSystemRights)SECURITY_INFORMATION.UNPROTECTED_SACL_SECURITY_INFORMATION) != 0 || (fileSystemRights & (FileSystemRights)SECURITY_INFORMATION.UNPROTECTED_DACL_SECURITY_INFORMATION) != 0 ? new PrivilegeEnabler(Privilege.Security) : null)

         using (var securityAttributes = new Security.NativeMethods.SecurityAttributes(fileSecurity))
         {
            var safeHandle = transaction == null || !NativeMethods.IsAtLeastWindowsVista

               // CreateFile() / CreateFileTransacted()
               // 2013-01-13: MSDNはLongPathの使用を確認しています。

               ? NativeMethods.CreateFile(pathLp, fileSystemRights, fileShare, securityAttributes, fileMode, attributes, IntPtr.Zero)

               : NativeMethods.CreateFileTransacted(pathLp, fileSystemRights, fileShare, securityAttributes, fileMode, attributes, IntPtr.Zero, transaction.SafeHandle, IntPtr.Zero, IntPtr.Zero);


            var lastError = Marshal.GetLastWin32Error();

            NativeMethods.CloseHandleAndPossiblyThrowException(safeHandle, lastError, isFolder, path, !continueOnException);


            if (isAppend)
            {
               var success = NativeMethods.SetFilePointerEx(safeHandle, 0, IntPtr.Zero, SeekOrigin.End);

               lastError = Marshal.GetLastWin32Error();

               if (!success)
               {
                  NativeMethods.CloseHandleAndPossiblyThrowException(safeHandle, lastError, isFolder, path, !continueOnException);

                  return null;
               }
            }

            return safeHandle;
         }
      }
   }
}
