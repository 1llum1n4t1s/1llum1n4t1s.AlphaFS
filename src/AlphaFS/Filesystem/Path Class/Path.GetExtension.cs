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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      #region .NET

      /// <summary>指定されたパス文字列の拡張子を返します。</summary>
      /// <returns>
      ///   <para>指定されたパスの拡張子（ピリオド "." を含む）、null、または <see cref="string.Empty"/>。</para>
      ///   <para><paramref name="path"/> が null の場合、このメソッドは null を返します。</para>
      ///   <para><paramref name="path"/> に拡張子情報がない場合、
      ///   このメソッドは <see cref="string.Empty"/> を返します。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">拡張子を取得するパス文字列。パスには <see cref="GetInvalidPathChars"/> で定義されている文字を含めることはできません。</param>
      [SecurityCritical]
      public static string GetExtension(string path)
      {
         return GetExtensionCore(path, !Utils.IsNullOrWhiteSpace(path));
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパス文字列の拡張子を返します。</summary>
      /// <returns>
      ///   <para>指定されたパスの拡張子（ピリオド "." を含む）、null、または <see cref="string.Empty"/>。</para>
      ///   <para><paramref name="path"/> が null の場合、このメソッドは null を返します。</para>
      ///   <para><paramref name="path"/> に拡張子情報がない場合、
      ///   このメソッドは <see cref="string.Empty"/> を返します。</para>
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <param name="path">拡張子を取得するパス文字列。パスには <see cref="GetInvalidPathChars"/> で定義されている文字を含めることはできません。</param>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="path"/> の無効なパス文字をチェックします。</param>
      [SecurityCritical]
      public static string GetExtension(string path, bool checkInvalidPathChars)
      {
         return GetExtensionCore(path, checkInvalidPathChars);
      }
   }
}
