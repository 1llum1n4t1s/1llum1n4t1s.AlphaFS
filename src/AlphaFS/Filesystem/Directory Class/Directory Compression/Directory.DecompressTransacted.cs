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
   public static partial class Directory
   {
      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <remarks>ルート項目のみを展開します（非再帰）。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, null, null, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <remarks>ルート項目のみを展開します（非再帰）。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, PathFormat pathFormat)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, null, null, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, options, null, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, PathFormat pathFormat)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, options, null, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <remarks>ルート項目のみを展開します（非再帰）。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationFilters filters)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, null, filters, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <remarks>ルート項目のみを展開します（非再帰）。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, null, filters, false, pathFormat);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, options, filters, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] NTFS圧縮されたディレクトリを展開します。</summary>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="DirectoryNotFoundException"/>
      /// <exception cref="IOException"/>
      /// <exception cref="NotSupportedException"/>
      /// <exception cref="UnauthorizedAccessException"/>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="path">展開するディレクトリを示すパス。</param>
      /// <param name="options">ディレクトリの列挙方法を指定する <see cref="DirectoryEnumerationOptions"/> フラグ。</param>
      /// <param name="filters">処理で使用するカスタムフィルタの指定。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void DecompressTransacted(KernelTransaction transaction, string path, DirectoryEnumerationOptions options, DirectoryEnumerationFilters filters, PathFormat pathFormat)
      {
         CompressDecompressCore(transaction, path, Path.WildcardStarMatchAll, options, filters, false, pathFormat);
      }
   }
}
