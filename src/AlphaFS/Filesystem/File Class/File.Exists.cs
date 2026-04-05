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
      /// <summary>指定されたファイルが存在するかどうかを判定します。</summary>
      /// <remarks>
      ///   <para>MSDN: .NET 3.5+: ファイルまたはディレクトリが存在するかどうかを
      ///   チェックする前に、<paramref name="path"/>パラメータの末尾のスペースが削除されます。</para>
      ///   <para>指定されたファイルが存在するかどうかを判定中にエラーが発生した場合、
      ///   Existsメソッドは<c>false</c>を返します。</para>
      ///   <para>これは、無効な文字を含むファイル名を渡す、
      ///   無効な文字が多すぎるか、ディスクの障害または欠落、あるいは呼び出し元にファイルの読み取り権限がない場合、
      ///   ファイル。</para>
      ///   <para>The Exists method should not be used for path validation,
      ///   this method merely checks pathで指定されたファイルが存在するかどうかをチェックするだけです。</para>
      ///   <para>無効なパスをExistsに渡すと、falseが返されます。</para>
      ///   <para>Existsメソッドの呼び出しからファイルに対する別の操作の実行までの間に、
      ///   Existsメソッドの呼び出しからファイルに対する別の操作(Deleteなど)の実行までの間に、</para>
      /// </remarks>
      /// <param name="path">チェックするファイル。</param>
      /// <returns>
      ///   Returns <c>true</c> if the caller has the required permissions and
      ///   <paramref name="path"/> contains the name of an existing file; otherwise,
      ///   <c>false</c>
      /// </returns>
      [SecurityCritical]
      public static bool Exists(string path)
      {
         return ExistsCore(null, false, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたファイルが存在するかどうかを判定します。</summary>
      /// <remarks>
      ///   <para>MSDN: .NET 3.5+: ファイルまたはディレクトリが存在するかどうかを
      ///   チェックする前に、<paramref name="path"/>パラメータの末尾のスペースが削除されます。</para>
      ///   <para>指定されたファイルが存在するかどうかを判定中にエラーが発生した場合、
      ///   Existsメソッドは<c>false</c>を返します。</para>
      ///   <para>これは、無効な文字を含むファイル名を渡す、
      ///   無効な文字が多すぎるなどの状況で発生する可能性があります。</para>
      ///   <para>a failing or missing disk, or if the caller does not have permission to read the
      ///   ファイル。</para>
      ///   <para>Existsメソッドはパスの検証に使用すべきではありません。このメソッドは単に
      ///   pathで指定されたファイルが存在するかどうかをチェックするだけです。</para>
      ///   <para>無効なパスをExistsに渡すと、falseが返されます。</para>
      ///   <para>Existsメソッドの呼び出しからファイルに対する別の操作の実行までの間に、
      ///   Existsメソッドの呼び出しからファイルに対する別の操作(Deleteなど)の実行
      ///   as Delete.</para>
      /// </remarks>
      /// <param name="path">チェックするファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns>
      ///   <para>Returns <c>true</c> if the caller has the required permissions and
      ///   <paramref name="path"/> contains the name of an existing file; otherwise,
      ///   <c>false</c></para>
      /// </returns>
      [SecurityCritical]
      public static bool Exists(string path, PathFormat pathFormat)
      {
         return ExistsCore(null, false, path, pathFormat);
      }
   }
}
