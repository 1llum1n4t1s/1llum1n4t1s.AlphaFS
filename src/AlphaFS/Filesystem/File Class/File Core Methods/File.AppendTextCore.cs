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
using System.IO;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] 既存のファイルにUTF-8エンコードされたテキストを追記する<see cref="StreamWriter"/>を作成します。指定されたファイルが存在しない場合は新しいファイルに書き込みます。</summary>
      /// <returns>指定されたファイルまたは新しいファイルにUTF-8エンコードされたテキストを追記するストリームライター。</returns>
      /// <exception cref="ArgumentException">path is a zero-length string, contains only white space, or contains one or more invalid characters as defined by InvalidPathChars.</exception>
      /// <exception cref="ArgumentNullException">path is null.</exception>
      /// <exception cref="DirectoryNotFoundException">The specified path is invalid (for example, the directory doesn�t exist or it is on an unmapped drive).</exception>
      /// <exception cref="NotSupportedException">path is in an invalid format.</exception>
      /// <exception cref="UnauthorizedAccessException">The caller does not have the required permission.</exception>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">追記するファイルのパス。</param>
      /// <param name="encoding">使用する文字<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      internal static StreamWriter AppendTextCore(KernelTransaction transaction, string path, Encoding encoding, PathFormat pathFormat)
      {
         var fs = OpenCore(transaction, path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None, ExtendedFileAttributes.Normal, null, null, pathFormat);

         try
         {
            fs.Seek(0, SeekOrigin.End);

            return new StreamWriter(fs, encoding);
         }
         catch (IOException)
         {
            fs.Close();

            throw;
         }
      }
   }
}
