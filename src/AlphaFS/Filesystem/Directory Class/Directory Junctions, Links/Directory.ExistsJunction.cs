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
      /// <summary>[AlphaFS] 指定されたパスがディスク上の既存のディレクトリジャンクションを参照しているかどうかを判定します。</summary>
      /// <returns>
      ///   <para>Returns <c>true</c> if <paramref name="junctionPath"/> refers to an existing directory junction.</para>
      ///   <para>Returns <c>false</c> if the directory junction does not exist or an error occurs when trying to determine if the specified file exists.</para>
      /// </returns>
      /// <para>&#160;</para>
      /// <remarks>
      ///   <para>The Exists method returns <c>false</c> if any error occurs while trying to determine if the specified file exists.</para>
      ///   <para>これは、無効な文字や文字数が多すぎるファイル名を渡すなど、例外が発生する状況で発生する可能性があります。</para>
      ///   <para>ディスクの障害や欠落、または呼び出し元にファイルの読み取り権限がない場合でも発生します。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">The path to test.</param>
      [SecurityCritical]
      public static bool ExistsJunction(string junctionPath)
      {
         return ExistsJunctionCore(null, null, junctionPath, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスがディスク上の既存のディレクトリジャンクションを参照しているかどうかを判定します。</summary>
      /// <returns>
      ///   <para>Returns <c>true</c> if <paramref name="junctionPath"/> refers to an existing directory junction.</para>
      ///   <para>Returns <c>false</c> if the directory junction does not exist or an error occurs when trying to determine if the specified file exists.</para>
      /// </returns>
      /// <para>&#160;</para>
      /// <remarks>
      ///   <para>The Exists method returns <c>false</c> if any error occurs while trying to determine if the specified file exists.</para>
      ///   <para>これは、無効な文字や文字数が多すぎるファイル名を渡すなど、例外が発生する状況で発生する可能性があります。</para>
      ///   <para>ディスクの障害や欠落、または呼び出し元にファイルの読み取り権限がない場合でも発生します。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">The path to test.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static bool ExistsJunction(string junctionPath, PathFormat pathFormat)
      {
         return ExistsJunctionCore(null, null, junctionPath, pathFormat);
      }
   }
}
