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
   public sealed partial class DirectoryInfo
   {
      /// <summary>[AlphaFS] 指定されたパスがディスク上の既存のディレクトリジャンクションを参照しているかどうかを判断します。</summary>
      /// <returns>
      ///   <para><paramref name="junctionPath"/> が既存のディレクトリジャンクションを参照している場合は <c>true</c> を返します。</para>
      ///   <para>ディレクトリジャンクションが存在しないか、指定されたファイルの存在を確認しようとしたときにエラーが発生した場合は <c>false</c> を返します。</para>
      /// </returns>
      /// <para>&#160;</para>
      /// <remarks>
      ///   <para>指定されたファイルの存在を確認しようとしたときにエラーが発生した場合、Exists メソッドは <c>false</c> を返します。</para>
      ///   <para>これは、無効な文字や文字数が多すぎるファイル名を渡した場合など、例外が発生する状況で起こる可能性があります。</para>
      ///   <para>また、ディスクの障害や欠落、またはファイルの読み取り権限がない場合にも発生します。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">テストするパス。</param>
      [SecurityCritical]
      public bool ExistsJunction(string junctionPath)
      {
         return Directory.ExistsJunctionCore(Transaction, null, junctionPath, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパスがディスク上の既存のディレクトリジャンクションを参照しているかどうかを判断します。</summary>
      /// <returns>
      ///   <para><paramref name="junctionPath"/> が既存のディレクトリジャンクションを参照している場合は <c>true</c> を返します。</para>
      ///   <para>ディレクトリジャンクションが存在しないか、指定されたファイルの存在を確認しようとしたときにエラーが発生した場合は <c>false</c> を返します。</para>
      /// </returns>
      /// <para>&#160;</para>
      /// <remarks>
      ///   <para>指定されたファイルの存在を確認しようとしたときにエラーが発生した場合、Exists メソッドは <c>false</c> を返します。</para>
      ///   <para>これは、無効な文字や文字数が多すぎるファイル名を渡した場合など、例外が発生する状況で起こる可能性があります。</para>
      ///   <para>また、ディスクの障害や欠落、またはファイルの読み取り権限がない場合にも発生します。</para>
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="junctionPath">テストするパス。</param>
      /// <param name="pathFormat">パスパラメーターの形式を示します。</param>
      [SecurityCritical]
      public bool ExistsJunction(string junctionPath, PathFormat pathFormat)
      {
         return Directory.ExistsJunctionCore(Transaction, null, junctionPath, pathFormat);
      }
   }
}
