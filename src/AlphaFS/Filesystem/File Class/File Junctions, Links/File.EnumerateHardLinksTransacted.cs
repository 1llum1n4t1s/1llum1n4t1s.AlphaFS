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
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region Obsolete

      /// <summary>[AlphaFS] 指定された<paramref name="path"/>へのすべてのハードリンクの列挙を作成します。</summary>
      /// <returns>指定された<paramref name="path"/>へのすべてのハードリンクの<see cref="string"/>の列挙可能なコレクション</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルの名前。</param>
      [Obsolete("Use EnumerateHardLinks method.")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Hardlinks")]
      [SecurityCritical]
      public static IEnumerable<string> EnumerateHardlinksTransacted(KernelTransaction transaction, string path)
      {
         return EnumerateHardLinksCore(transaction, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された<paramref name="path"/>へのすべてのハードリンクの列挙を作成します。</summary>
      /// <returns>指定された<paramref name="path"/>へのすべてのハードリンクの<see cref="string"/>の列挙可能なコレクション</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [Obsolete("Use EnumerateHardLinks method.")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Hardlinks")]
      [SecurityCritical]
      public static IEnumerable<string> EnumerateHardlinksTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return EnumerateHardLinksCore(transaction, path, pathFormat);
      }

      #endregion // Obsolete


      /// <summary>[AlphaFS] 指定された<paramref name="path"/>へのすべてのハードリンクの列挙を作成します。</summary>
      /// <returns>指定された<paramref name="path"/>へのすべてのハードリンクの<see cref="string"/>の列挙可能なコレクション</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルの名前。</param>
      [SecurityCritical]
#pragma warning disable CS3005 // Identifier differing only in case is not CLS-compliant
      public static IEnumerable<string> EnumerateHardLinksTransacted(KernelTransaction transaction, string path)
#pragma warning restore CS3005 // Identifier differing only in case is not CLS-compliant
      {
         return EnumerateHardLinksCore(transaction, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された<paramref name="path"/>へのすべてのハードリンクの列挙を作成します。</summary>
      /// <returns>指定された<paramref name="path"/>へのすべてのハードリンクの<see cref="string"/>の列挙可能なコレクション</returns>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">ファイルの名前。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
#pragma warning disable CS3005 // Identifier differing only in case is not CLS-compliant
      public static IEnumerable<string> EnumerateHardLinksTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
#pragma warning restore CS3005 // Identifier differing only in case is not CLS-compliant
      {
         return EnumerateHardLinksCore(transaction, path, pathFormat);
      }
   }
}
