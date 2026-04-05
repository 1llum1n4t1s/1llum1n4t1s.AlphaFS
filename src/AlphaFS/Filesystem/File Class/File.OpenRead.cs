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

using System.IO;
using System.Security;
using FileStream = System.IO.FileStream;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>既存のファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <returns>指定されたパス上の読み取り専用<see cref="FileStream"/>。</returns>
      /// <remarks>
      ///   このメソッドは<see cref="FileStream"/>(string, FileMode, FileAccess, FileShare)コンストラクタオーバーロードで、
      ///   <see cref="FileMode"/> value of Open, a <see cref="FileAccess"/> value of Read and a <see cref="FileShare"/> value of Read.
      /// </remarks>
      [SecurityCritical]
      public static FileStream OpenRead(string path)
      {
         return Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 既存のファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたパス上の読み取り専用<see cref="FileStream"/>。</returns>
      /// <remarks>
      ///   このメソッドは<see cref="FileStream"/>(string, FileMode, FileAccess, FileShare)コンストラクタオーバーロードで、
      ///   <see cref="FileMode"/> value of Open, a <see cref="FileAccess"/> value of Read and a <see cref="FileShare"/> value of Read.
      /// </remarks>
      [SecurityCritical]
      public static FileStream OpenRead(string path, PathFormat pathFormat)
      {
         return Open(path, FileMode.Open, FileAccess.Read, FileShare.Read, pathFormat);
      }
   }
}
