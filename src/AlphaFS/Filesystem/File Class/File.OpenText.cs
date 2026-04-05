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
using System.IO;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>既存のUTF-8エンコードされたテキストファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <returns>指定されたパス上の<see cref="StreamReader"/>。</returns>
      /// <remarks>このメソッドは<see cref="StreamReader"/>(String)コンストラクタオーバーロードと同等です。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamReader OpenText(string path)
      {
         return new StreamReader(OpenRead(path), NativeMethods.DefaultFileEncoding);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 既存のUTF-8エンコードされたテキストファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたパス上の<see cref="StreamReader"/>。</returns>
      /// <remarks>このメソッドは<see cref="StreamReader"/>(String)コンストラクタオーバーロードと同等です。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamReader OpenText(string path, PathFormat pathFormat)
      {
         return new StreamReader(OpenRead(path, pathFormat), NativeMethods.DefaultFileEncoding);
      }


      /// <summary>[AlphaFS] 既存の<see cref="Encoding"/>エンコードされたテキストファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>指定されたパス上の<see cref="StreamReader"/>。</returns>
      /// <remarks>このメソッドは<see cref="StreamReader"/>(String)コンストラクタオーバーロードと同等です。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamReader OpenText(string path, Encoding encoding, PathFormat pathFormat)
      {
         return new StreamReader(OpenRead(path, pathFormat), encoding);
      }


      /// <summary>[AlphaFS] 既存の<see cref="Encoding"/>エンコードされたテキストファイルを読み取り用に開きます。</summary>
      /// <param name="path">読み取り用に開くファイル。</param>
      /// <param name="encoding">ファイルの内容に適用される<see cref="Encoding"/>。</param>
      /// <returns>指定されたパス上の<see cref="StreamReader"/>。</returns>
      /// <remarks>このメソッドは<see cref="StreamReader"/>(String)コンストラクタオーバーロードと同等です。</remarks>
      [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
      [SecurityCritical]
      public static StreamReader OpenText(string path, Encoding encoding)
      {
         return new StreamReader(OpenRead(path), encoding);
      }
   }
}
