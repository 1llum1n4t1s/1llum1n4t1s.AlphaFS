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
      /// <summary>[AlphaFS] 指定されたパス文字列からルート情報を除いたディレクトリ情報を返します。例: "C:\Windows\system32" は "Windows" を返します。</summary>
      /// <returns>ファイル名部分とルート情報（存在する場合）を除いた <paramref name="path"/>。<paramref name="path"/> が <c>null</c> の場合、またはルートを示す場合（"\"、"C:"、"\\server\share" など）は <c>null</c> を返します。</returns>
      /// <param name="path">パス。</param>
      [SecurityCritical]
      public static string GetDirectoryNameWithoutRoot(string path)
      {
         return GetDirectoryNameWithoutRootCore(null, path, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたパス文字列からルート情報を除いたディレクトリ情報を返します。例: "C:\Windows\system32" は "Windows" を返します。</summary>
      /// <returns>ファイル名部分とルート情報（存在する場合）を除いた <paramref name="path"/>。<paramref name="path"/> が <c>null</c> の場合、またはルートを示す場合（"\"、"C:"、"\\server\share" など）は <c>null</c> を返します。</returns>
      /// <param name="path">パス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static string GetDirectoryNameWithoutRoot(string path, PathFormat pathFormat)
      {
         return GetDirectoryNameWithoutRootCore(null, path, pathFormat);
      }
   }
}
