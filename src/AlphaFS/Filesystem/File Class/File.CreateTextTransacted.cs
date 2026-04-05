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

using System.Diagnostics.CodeAnalysis;
using System.Security;
using System.Text;
using StreamWriter = System.IO.StreamWriter;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] UTF-8エンコードされたテキストの書き込み用にファイルを作成または開きます。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込み用に開くファイル。</param>
      /// <returns>UTF-8エンコーディングを使用して指定されたファイルに書き込むStreamWriter。</returns>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamWriter CreateTextTransacted(KernelTransaction transaction, string path)
      {
         return CreateTextCore(transaction, path, NativeMethods.DefaultFileEncoding, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] <see cref="Encoding"/>エンコードされたテキストの書き込み用にファイルを作成または開きます。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込み用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>UTF-8エンコーディングを使用して指定されたファイルに書き込むStreamWriter。</returns>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamWriter CreateTextTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return CreateTextCore(transaction, path, NativeMethods.DefaultFileEncoding, pathFormat);
      }


      /// <summary>[AlphaFS] <see cref="Encoding"/>エンコードされたテキストの書き込み用にファイルを作成または開きます。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込み用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用されるエンコーディング。</param>
      /// <returns>UTF-8エンコーディングを使用して指定されたファイルに書き込むStreamWriter。</returns>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamWriter CreateTextTransacted(KernelTransaction transaction, string path, Encoding encoding)
      {
         return CreateTextCore(transaction, path, encoding, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] <see cref="Encoding"/>エンコードされたテキストの書き込み用にファイルを作成または開きます。</summary>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">書き込み用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用されるエンコーディング。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>UTF-8エンコーディングを使用して指定されたファイルに書き込むStreamWriter。</returns>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamWriter CreateTextTransacted(KernelTransaction transaction, string path, Encoding encoding, PathFormat pathFormat)
      {
         return CreateTextCore(transaction, path, encoding, pathFormat);
      }
   }
}
