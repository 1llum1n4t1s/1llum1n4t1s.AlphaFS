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

using System.Globalization;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>バイナリファイルを開き、ファイルの内容をバイト配列に読み込み、ファイルを閉じます。</summary>
      /// <exception cref="IOException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ファイルの内容を含むバイト配列。</returns>
      [SecurityCritical]
      internal static byte[] ReadAllBytesCore(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         byte[] buffer;

         using var fs = OpenReadTransacted(transaction, path, pathFormat);
         var offset = 0;
         var length = fs.Length;

         if (length > int.MaxValue)
         {
            throw new IOException(string.Format(CultureInfo.InvariantCulture, "File larger than 2GB: [{0}]", path));
         }

         var count = (int) length;
         buffer = new byte[count];

         while (count > 0)
         {
            var n = fs.Read(buffer, offset, count);
            if (n == 0)
            {
               throw new IOException("UNEXPECTED end of file found");
            }
            offset += n;
            count -= n;
         }

         return buffer;
      }
   }
}
