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
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path)
      {
         return CreateSubdirectoryCore(path, null, null, false);
      }


      /// <summary>指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <param name="directorySecurity">適用する <see cref="DirectorySecurity"/> セキュリティ。</param>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path, DirectorySecurity directorySecurity)
      {
         return CreateSubdirectoryCore(path, null, directorySecurity, false);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path, bool compress)
      {
         return CreateSubdirectoryCore(path, null, null, compress);
      }


      /// <summary>[AlphaFS] 指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <param name="templatePath">新しいディレクトリの作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path, string templatePath, bool compress)
      {
         return CreateSubdirectoryCore(path, templatePath, null, compress);
      }


      /// <summary>[AlphaFS] 指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <param name="directorySecurity">適用する <see cref="DirectorySecurity"/> セキュリティ。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path, DirectorySecurity directorySecurity, bool compress)
      {
         return CreateSubdirectoryCore(path, null, directorySecurity, compress);
      }


      /// <summary>[AlphaFS] 指定されたパスにサブディレクトリを作成します。指定されたパスは、この <see cref="DirectoryInfo"/> クラスのインスタンスに対する相対パスにできます。</summary>
      /// <param name="templatePath">新しいディレクトリの作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="path">指定されたパス。別のディスクボリュームにすることはできません。</param>
      /// <param name="compress"><c>true</c> の場合、NTFS 圧縮を使用してディレクトリを圧縮します。</param>
      /// <param name="directorySecurity">適用する <see cref="DirectorySecurity"/> セキュリティ。</param>
      /// <returns><paramref name="path"/> で指定された最後のディレクトリ。</returns>
      /// <remarks>
      /// パスの一部が無効でない限り、パスに指定されたすべてのディレクトリが作成されます。
      /// path パラメーターはディレクトリパスを指定するもので、ファイルパスではありません。
      /// サブディレクトリが既に存在する場合、このメソッドは何もしません。
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public DirectoryInfo CreateSubdirectory(string path, string templatePath, DirectorySecurity directorySecurity, bool compress)
      {
         return CreateSubdirectoryCore(path, templatePath, directorySecurity, compress);
      }
   }
}
