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
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      /// <summary>指定されたパスにサブディレクトリを作成します。指定されたパスは、この DirectoryInfo クラスのインスタンスに対する相対パスにできます。</summary>
      /// <returns>パスで指定された最後のディレクトリを <see cref="DirectoryInfo"/> オブジェクトとして返します。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      /// <param name="path">指定されたパス。別のディスクボリュームまたは UNC (Universal Naming Convention) 名にすることはできません。</param>
      /// <param name="templatePath">新しいディレクトリの作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="directorySecurity">適用する <see cref="DirectorySecurity"/> セキュリティ。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      [SecurityCritical]
      private DirectoryInfo CreateSubdirectoryCore(string path, string templatePath, ObjectSecurity directorySecurity, bool compress)
      {
         var pathLp = Path.CombineCore(false, LongFullName, path);

         var templatePathLp = null == templatePath ? null : Path.GetExtendedLengthPathCore(Transaction, templatePath, PathFormat.RelativePath, GetFullPathOptions.TrimEnd | GetFullPathOptions.RemoveTrailingDirectorySeparator);


         if (string.Compare(LongFullName, 0, pathLp, 0, LongFullName.Length, StringComparison.OrdinalIgnoreCase) != 0)

         {
            throw new ArgumentException(Resources.Invalid_Subpath, "path");
         }


         return Directory.CreateDirectoryCore(false, Transaction, pathLp, templatePathLp, directorySecurity, compress, PathFormat.LongFullPath);
      }
   }
}
