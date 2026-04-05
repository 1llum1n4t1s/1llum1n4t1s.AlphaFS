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
using System.Collections.Generic;
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Directory
   {
      /// <summary>[AlphaFS] シンボリックリンクやマウントポイントをたどらずに、特定のディレクトリのプロパティを取得します。
      ///   <para>Properties include aggregated info from <see cref="FileAttributes"/> of each encountered file system object, plus additional ones: Total, File, Size and Error.</para>
      ///   <para><b>Total:</b> 列挙されたオブジェクトの総数。</para>
      ///   <para><b>File:</b> is the total number of files. File is considered when object is neither <see cref="FileAttributes.Directory"/> nor <see cref="FileAttributes.ReparsePoint"/>.</para>
      ///   <para><b>Size:</b> 列挙されたオブジェクトの合計サイズ。</para>
      ///   <para><b>Error:</b> 列挙中に発生したエラーの総数。</para>
      /// </summary>
      /// <returns>上記のキーをそれぞれの集計値にマッピングしたディクショナリ。</returns>
      /// <remarks><b>Directory:</b> is an object which has <see cref="FileAttributes.Directory"/> attribute without <see cref="FileAttributes.ReparsePoint"/> one.</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The target directory.</param>
      [SecurityCritical]
      public static Dictionary<string, long> GetPropertiesTransacted(KernelTransaction transaction, string path)
      {
         return GetPropertiesCore(transaction, path, null, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] シンボリックリンクやマウントポイントをたどらずに、特定のディレクトリのプロパティを取得します。
      ///   <para>Properties include aggregated info from <see cref="FileAttributes"/> of each encountered file system object, plus additional ones: Total, File, Size and Error.</para>
      ///   <para><b>Total:</b> 列挙されたオブジェクトの総数。</para>
      ///   <para><b>File:</b> is the total number of files. File is considered when object is neither <see cref="FileAttributes.Directory"/> nor <see cref="FileAttributes.ReparsePoint"/>.</para>
      ///   <para><b>Size:</b> 列挙されたオブジェクトの合計サイズ。</para>
      ///   <para><b>Error:</b> 列挙中に発生したエラーの総数。</para>
      /// </summary>
      /// <returns>上記のキーをそれぞれの集計値にマッピングしたディクショナリ。</returns>
      /// <remarks><b>Directory:</b> is an object which has <see cref="FileAttributes.Directory"/> attribute without <see cref="FileAttributes.ReparsePoint"/> one.</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The target directory.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static Dictionary<string, long> GetPropertiesTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         return GetPropertiesCore(transaction, path, null, pathFormat);
      }


      /// <summary>[AlphaFS] シンボリックリンクやマウントポイントをたどらずに、特定のディレクトリのプロパティを取得します。
      ///   <para>Properties include aggregated info from <see cref="FileAttributes"/> of each encountered file system object, plus additional ones: Total, File, Size and Error.</para>
      ///   <para><b>Total:</b> 列挙されたオブジェクトの総数。</para>
      ///   <para><b>File:</b> is the total number of files. File is considered when object is neither <see cref="FileAttributes.Directory"/> nor <see cref="FileAttributes.ReparsePoint"/>.</para>
      ///   <para><b>Size:</b> 列挙されたオブジェクトの合計サイズ。</para>
      ///   <para><b>Error:</b> 列挙中に発生したエラーの総数。</para>
      /// </summary>
      /// <returns>上記のキーをそれぞれの集計値にマッピングしたディクショナリ。</returns>
      /// <remarks><b>Directory:</b> is an object which has <see cref="FileAttributes.Directory"/> attribute without <see cref="FileAttributes.ReparsePoint"/> one.</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The target directory.</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public static Dictionary<string, long> GetPropertiesTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options)
      {
         return GetPropertiesCore(transaction, path, options, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] シンボリックリンクやマウントポイントをたどらずに、特定のディレクトリのプロパティを取得します。
      ///   <para>Properties include aggregated info from <see cref="FileAttributes"/> of each encountered file system object, plus additional ones: Total, File, Size and Error.</para>
      ///   <para><b>Total:</b> 列挙されたオブジェクトの総数。</para>
      ///   <para><b>File:</b> is the total number of files. File is considered when object is neither <see cref="FileAttributes.Directory"/> nor <see cref="FileAttributes.ReparsePoint"/>.</para>
      ///   <para><b>Size:</b> 列挙されたオブジェクトの合計サイズ。</para>
      ///   <para><b>Error:</b> 列挙中に発生したエラーの総数。</para>
      /// </summary>
      /// <returns>上記のキーをそれぞれの集計値にマッピングしたディクショナリ。</returns>
      /// <remarks><b>Directory:</b> is an object which has <see cref="FileAttributes.Directory"/> attribute without <see cref="FileAttributes.ReparsePoint"/> one.</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">The target directory.</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static Dictionary<string, long> GetPropertiesTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         return GetPropertiesCore(transaction, path, options, pathFormat);
      }
   }
}
