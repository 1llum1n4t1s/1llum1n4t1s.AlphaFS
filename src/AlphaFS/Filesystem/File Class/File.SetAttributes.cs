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
   public static partial class File
   {
      #region .NET

      /// <summary>指定されたパス上のファイルまたはディレクトリに指定された<see cref="FileAttributes"/>を設定します。</summary>
      /// <remarks>
      ///   <see cref="FileAttributes.Hidden"/>や<see cref="FileAttributes.ReadOnly"/>などの特定のファイル属性は組み合わせることができます。
      ///   <see cref="FileAttributes.Normal"/>などの他の属性は単独で使用する必要があります。
      /// </remarks>
      /// <remarks>
      ///   このメソッドを使用してFileオブジェクトの<see cref="FileAttributes.Compressed"/>ステータスを変更することはできません。
      /// </remarks>
      /// <exception cref="ArgumentException">path is empty, contains only white spaces, contains invalid characters, or the file attribute is invalid.</exception>
      /// <exception cref="DirectoryNotFoundException">The specified path is invalid, (for example, it is on an unmapped drive).</exception>
      /// <exception cref="FileNotFoundException">The file cannot be found.</exception>
      /// <exception cref="NotSupportedException">path is in an invalid format.</exception>
      /// <exception cref="UnauthorizedAccessException">path specified a file that is read-only. -or- This operation is not supported on the current platform. -or- path specified a directory. -or- The caller does not have the required permission.</exception>
      /// <param name="path">ファイルまたはディレクトリへのパス。</param>
      /// <param name="fileAttributes">列挙値のビット単位の組み合わせ。</param>
      /// <overloads>指定されたパス上のファイルまたはディレクトリに指定された<see cref="FileAttributes"/>を設定します。</overloads>
      [SecurityCritical]
      public static void SetAttributes(string path, FileAttributes fileAttributes)
      {
         SetAttributesCore(null, false, path, fileAttributes, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたパス上のファイルまたはディレクトリに指定された<see cref="FileAttributes"/>を設定します。</summary>
      /// <remarks>
      ///   <see cref="FileAttributes.Hidden"/>や<see cref="FileAttributes.ReadOnly"/>などの特定のファイル属性は組み合わせることができます。
      ///   <see cref="FileAttributes.Normal"/>などの他の属性は単独で使用する必要があります。
      /// </remarks>
      /// <remarks>
      ///   このメソッドを使用してFileオブジェクトの<see cref="FileAttributes.Compressed"/>ステータスを変更することはできません。
      /// </remarks>
      /// <exception cref="ArgumentException">path is empty, contains only white spaces, contains invalid characters, or the file attribute is invalid.</exception>
      /// <exception cref="DirectoryNotFoundException">The specified path is invalid, (for example, it is on an unmapped drive).</exception>
      /// <exception cref="FileNotFoundException">The file cannot be found.</exception>
      /// <exception cref="NotSupportedException">path is in an invalid format.</exception>
      /// <exception cref="UnauthorizedAccessException">path specified a file that is read-only. -or- This operation is not supported on the current platform. -or- path specified a directory. -or- The caller does not have the required permission.</exception>
      /// <param name="path">ファイルまたはディレクトリへのパス。</param>
      /// <param name="fileAttributes">列挙値のビット単位の組み合わせ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SecurityCritical]
      public static void SetAttributes(string path, FileAttributes fileAttributes, PathFormat pathFormat)
      {
         SetAttributesCore(null, false, path, fileAttributes, pathFormat);
      }
   }
}
