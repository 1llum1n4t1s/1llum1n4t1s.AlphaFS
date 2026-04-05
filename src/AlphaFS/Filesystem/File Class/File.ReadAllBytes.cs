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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>バイナリファイルを開き、ファイルの内容をバイト配列に読み込み、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <returns>ファイルの内容を含むバイト配列。</returns>
      [SecurityCritical]
      public static byte[] ReadAllBytes(string path)
      {
         return ReadAllBytesCore(null, path, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] バイナリファイルを開き、ファイルの内容をバイト配列に読み込み、ファイルを閉じます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>ファイルの内容を含むバイト配列。</returns>
      [SecurityCritical]
      public static byte[] ReadAllBytes(string path, PathFormat pathFormat)
      {
         return ReadAllBytesCore(null, path, pathFormat);
      }
   }
}
