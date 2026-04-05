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
   public static partial class Path
   {
      /// <summary>[AlphaFS] 指定された <paramref name="path"/> のディレクトリ情報を末尾に <see cref="DirectorySeparatorChar"/> 文字を付加して返します。</summary>
      /// <returns>
      ///   <para>指定された <paramref name="path"/> の末尾に <see cref="DirectorySeparatorChar"/> 文字を付加したディレクトリ情報。</para>
      ///   <para><paramref name="path"/> が <c>null</c> の場合、またはルートを示す場合（"\"、"C:"、"\\server\share" など）は <c>null</c>。</para>
      /// </returns>
      /// <remarks>このメソッドは Path.GetDirectoryName() + Path.AddTrailingDirectorySeparator() を呼び出すのと同様です。</remarks>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">パス。</param>
      [SecurityCritical]
      public static string GetSuffixedDirectoryNameTransacted(KernelTransaction transaction, string path)
      {
         return GetSuffixedDirectoryNameCore(transaction, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された <paramref name="path"/> のディレクトリ情報を末尾に <see cref="DirectorySeparatorChar"/> 文字を付加して返します。</summary>
      /// <returns>
      ///   <para>指定された <paramref name="path"/> の末尾に <see cref="DirectorySeparatorChar"/> 文字を付加したディレクトリ情報。</para>
      ///   <para><paramref name="path"/> が <c>null</c> の場合、またはルートを示す場合（"\"、"C:"、"\\server\share" など）は <c>null</c>。</para>
      /// </returns>
      /// <remarks>このメソッドは Path.GetDirectoryName() + Path.AddTrailingDirectorySeparator() を呼び出すのと同様です。</remarks>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">パス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static string GetSuffixedDirectoryNameTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return GetSuffixedDirectoryNameCore(transaction, path, pathFormat);
      }
   }
}
