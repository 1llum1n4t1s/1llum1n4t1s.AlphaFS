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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>[AlphaFS] 指定されたディレクトリとその中のファイルのNTFS圧縮を無効にします。</summary>
      /// <remarks>このメソッドはディレクトリの圧縮属性を無効にします。ディレクトリの現在の内容は展開されません。ただし、新しく作成されるファイルとディレクトリは非圧縮になります。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリへのパス。</param>
      [SecurityCritical]
      public static void DisableCompressionTransacted(KernelTransaction transaction, string path)
      {
         Device.ToggleCompressionCore(transaction, true, path, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたディレクトリとその中のファイルのNTFS圧縮を無効にします。</summary>
      /// <remarks>このメソッドはディレクトリの圧縮属性を無効にします。ディレクトリの現在の内容は展開されません。ただし、新しく作成されるファイルとディレクトリは非圧縮になります。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <param name="path">展開するディレクトリへのパス。</param>
      [SecurityCritical]
      public static void DisableCompressionTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         Device.ToggleCompressionCore(transaction, true, path, false, pathFormat);
      }
   }
}
