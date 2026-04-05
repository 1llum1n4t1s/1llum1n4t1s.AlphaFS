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
   public static partial class Directory
   {
      /// <summary>絶対パスと相対パスの両方を含む、指定されたパスの親ディレクトリを取得します。</summary>
      /// <returns>親ディレクトリ。<paramref name="path"/> がルートディレクトリ（UNCサーバーまたは共有名のルートを含む）の場合は <c>null</c>。</returns>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">親ディレクトリを取得するパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      internal static DirectoryInfo GetParentCore(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         var pathLp = pathFormat == PathFormat.LongFullPath ? path : Path.GetExtendedLengthPathCore(transaction, path, pathFormat, GetFullPathOptions.CheckInvalidPathChars);

         pathLp = Path.GetRegularPathCore(pathLp, GetFullPathOptions.None, false);

         var dirName = Path.GetDirectoryName(pathLp, false);

         return !Utils.IsNullOrWhiteSpace(dirName) ? new DirectoryInfo(transaction, dirName, PathFormat.RelativePath) : null;
      }
   }
}
