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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>[AlphaFS] <paramref name="path"/> が "C:"、"D:" のような論理ドライブ形式であるかどうかをチェックします。</summary>
      /// <returns><paramref name="path"/> が "C:"、"D:" のような論理ドライブ形式の場合に true。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">チェックする絶対パス。</param>
      public static bool IsLogicalDrive(string path)
      {
         return IsLogicalDriveCore(path, false, PathFormat.FullPath);
      }


      /// <summary>[AlphaFS] <paramref name="path"/> が "C:"、"D:" のような論理ドライブ形式であるかどうかをチェックします。</summary>
      /// <returns><paramref name="path"/> が "C:"、"D:" のような論理ドライブ形式の場合に true。</returns>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">チェックする絶対パス。</param>
      /// <param name="isRegularPath"><c>true</c> はパスが既に通常のパスであることを示します。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      internal static bool IsLogicalDriveCore(string path, bool isRegularPath, PathFormat pathFormat)
      {
         if (pathFormat != PathFormat.LongFullPath)
         {
            if (Utils.IsNullOrWhiteSpace(path))
            {
               throw new ArgumentNullException("path");
            }

            CheckSupportedPathFormat(path, true, true);
         }


         if (!isRegularPath)
         {
            path = GetRegularPathCore(path, GetFullPathOptions.None, false);
         }

         var regularPath = path.StartsWith(LogicalDrivePrefix, StringComparison.OrdinalIgnoreCase) ? path.Substring(LogicalDrivePrefix.Length) : path;
         
         // 必要なのは先頭 1 文字だけなので、パス全体を大文字化した文字列を確保しない。
         // ここは Path.GetLongPathCore 経由でほぼ全ての公開 API が通るホットパスで、
         // 拡張長パス (最大 32,700 文字) では 1 呼び出しあたり数十 KB のアロケーションになっていた。
         var c = char.ToUpperInvariant(regularPath[0]);

         // char.IsLetter() は誤解を招く可能性があるため使用しない。有効なドライブ文字は A-Z のみ。

         return regularPath[1] == VolumeSeparatorChar && c >= 'A' && c <= 'Z';
      }
   }
}
