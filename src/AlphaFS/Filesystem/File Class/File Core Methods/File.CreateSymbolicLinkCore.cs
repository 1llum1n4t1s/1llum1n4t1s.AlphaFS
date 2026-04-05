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
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] トランザクション操作として、ファイルまたはディレクトリへのシンボリックリンク(CMDコマンド"MKLINK"と同様)を作成します。</summary>
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>シンボリックリンクは存在しないターゲットを指すことができます。</para>
      /// <para>シンボリックリンクを作成するとき、オペレーティングシステムはターゲットが存在するかどうかをチェックしません。</para>
      /// <para>シンボリックリンクはリパースポイントです。</para>
      /// <para>特定のパスで許可されるリパースポイント(したがってシンボリックリンク)は最大31個です。</para>
      /// <para>このメソッドを昇格された状態で実行するには、<see cref="Security.Privilege.CreateSymbolicLink"/>を参照してください。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="symlinkFileName">作成するシンボリックリンクのターゲット名。</param>
      /// <param name="targetFileName">作成するシンボリックリンク。</param>
      /// <param name="targetType">リンクターゲット<paramref name="targetFileName"/>がファイルかディレクトリかを示します。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static void CreateSymbolicLinkCore(KernelTransaction transaction, string symlinkFileName, string targetFileName, SymbolicLinkTarget targetType, PathFormat pathFormat)
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }


         if (pathFormat != PathFormat.LongFullPath)
         {
            const GetFullPathOptions options = GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck;

            symlinkFileName = Path.GetExtendedLengthPathCore(transaction, symlinkFileName, pathFormat, options);
            targetFileName = Path.GetExtendedLengthPathCore(transaction, targetFileName, pathFormat, options);
         }


         // ロングパス表記は使用しない。作成時に空になるため。
         targetFileName = Path.GetRegularPathCore(targetFileName, GetFullPathOptions.None, false);


         if (targetType == SymbolicLinkTarget.Directory)
         {
            ThrowIOExceptionIfFsoExist(transaction, false, targetFileName, pathFormat);
            ThrowIOExceptionIfFsoExist(transaction, false, symlinkFileName, pathFormat);
         }

         else
         {
            ThrowIOExceptionIfFsoExist(transaction, true, targetFileName, pathFormat);
            ThrowIOExceptionIfFsoExist(transaction, true, symlinkFileName, pathFormat);
         }


         var success = null == transaction

            // CreateSymbolicLink() / CreateSymbolicLinkTransacted()
            // 2017-05-30: CreateSymbolicLink() MSDNはLongPathの使用を確認: Windows 10 バージョン1607以降
            // 2015-07-17: この関数はロングパスをサポートしていません。
            // 2014-02-14: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。
            
            ? NativeMethods.CreateSymbolicLink(symlinkFileName, targetFileName, targetType)
            : NativeMethods.CreateSymbolicLinkTransacted(symlinkFileName, targetFileName, targetType, transaction.SafeHandle);


         var lastError = (uint) Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError, targetFileName, symlinkFileName);
         }
      }
   }
}
