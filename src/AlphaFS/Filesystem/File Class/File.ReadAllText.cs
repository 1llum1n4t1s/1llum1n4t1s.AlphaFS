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

using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>テキストファイルを開き、ファイルのすべての行を読み取り、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <returns>ファイルのすべての行。</returns>
      [SecurityCritical]
      public static string ReadAllText(string path)
      {
         return ReadAllTextCore(null, path, NativeMethods.DefaultFileEncoding, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] テキストファイルを開き、ファイルのすべての行を読み取り、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ファイルのすべての行。</returns>
      [SecurityCritical]
      public static string ReadAllText(string path, PathFormat pathFormat)
      {
         return ReadAllTextCore(null, path, NativeMethods.DefaultFileEncoding, pathFormat);
      }


      /// <summary>ファイルを開き、指定されたエンコーディングでファイルのすべての行を読み取り、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      /// <returns>ファイルのすべての行。</returns>
      [SecurityCritical]
      public static string ReadAllText(string path, Encoding encoding)
      {
         return ReadAllTextCore(null, path, encoding, PathFormat.RelativePath);
      }
      

      /// <summary>[AlphaFS] ファイルを開き、指定されたエンコーディングでファイルのすべての行を読み取り、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ファイルのすべての行。</returns>
      [SecurityCritical]
      public static string ReadAllText(string path, Encoding encoding, PathFormat pathFormat)
      {
         return ReadAllTextCore(null, path, encoding, pathFormat);
      }
   }
}
