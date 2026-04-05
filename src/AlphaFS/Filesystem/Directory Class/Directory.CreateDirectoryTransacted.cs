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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>[AlphaFS] 指定されたパス内のすべてのディレクトリとサブディレクトリを、既に存在しない限り作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path)
      {
         return CreateDirectoryCore(false, transaction, path, null, null, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, null, null, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, bool compress)
      {
         return CreateDirectoryCore(false, transaction, path, null, null, compress, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, bool compress, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, null, null, compress, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたパス内のすべてのディレクトリを、既に存在しない限り、指定されたWindowsセキュリティを適用して作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, DirectorySecurity directorySecurity)
      {
         return CreateDirectoryCore(false, transaction, path, null, directorySecurity, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, DirectorySecurity directorySecurity, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, null, directorySecurity, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, DirectorySecurity directorySecurity, bool compress)
      {
         return CreateDirectoryCore(false, transaction, path, null, directorySecurity, compress, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたWindowsセキュリティを適用して、指定されたパス内のすべてのディレクトリを作成します。</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, DirectorySecurity directorySecurity, bool compress, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, null, directorySecurity, compress, pathFormat);
      }


      /// <summary>[AlphaFS] Creates a new directory, with the attributes of a specified template directory.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, null, false, PathFormat.RelativePath);
      }

      /// <summary>[AlphaFS] Creates a new directory, with the attributes of a specified template directory.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, null, false, pathFormat);
      }


      /// <summary>[AlphaFS] Creates a new directory, with the attributes of a specified template directory.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, bool compress)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, null, compress, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] Creates a new directory, with the attributes of a specified template directory.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, bool compress, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, null, compress, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたテンプレートディレクトリの指定されたパス内にすべてのディレクトリを作成します and applies the specified Windows security.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, DirectorySecurity directorySecurity)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, directorySecurity, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたテンプレートディレクトリの指定されたパス内にすべてのディレクトリを作成します and applies the specified Windows security.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, DirectorySecurity directorySecurity, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, directorySecurity, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたテンプレートディレクトリの指定されたパス内にすべてのディレクトリを作成します and applies the specified Windows security.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, DirectorySecurity directorySecurity, bool compress)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, directorySecurity, compress, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたテンプレートディレクトリの指定されたパス内にすべてのディレクトリを作成します and applies the specified Windows security.</summary>
      /// <returns>指定されたパスのディレクトリを表すオブジェクト。指定されたパスにディレクトリが既に存在するかどうかに関係なく、このオブジェクトが返されます。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">作成するディレクトリ。</param>
      /// <param name="templatePath">新しいディレクトリ作成時にテンプレートとして使用するディレクトリのパス。</param>
      /// <param name="directorySecurity">ディレクトリに適用するアクセス制御。</param>
      /// <param name="compress">When <c>true</c> compresses the directory using NTFS compression.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static DirectoryInfo CreateDirectoryTransacted(KernelTransaction transaction, string path, string templatePath, DirectorySecurity directorySecurity, bool compress, PathFormat pathFormat)
      {
         return CreateDirectoryCore(false, transaction, path, templatePath, directorySecurity, compress, pathFormat);
      }
   }
}
