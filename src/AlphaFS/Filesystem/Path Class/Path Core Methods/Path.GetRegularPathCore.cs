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

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>長いパスから通常のパスを取得します。</summary>
      /// <returns>
      ///   長い <paramref name="path"/> の通常形式。
      ///   例: "\\?\C:\Temp\file.txt" は "C:\Temp\file.txt" に、"\\?\UNC\Server\share\file.txt" は "\\Server\share\file.txt" に変換されます。
      /// </returns>
      /// <remarks>
      ///   MSDN: String.TrimEnd Method notes to Callers: http://msdn.microsoft.com/en-us/library/system.string.trimend%28v=vs.110%29.aspx
      /// </remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">パス。</param>
      /// <param name="options">完全パス取得を制御するオプション。</param>
      /// <param name="allowEmpty"><c>false</c> の場合、<see cref="ArgumentException"/> をスローします。</param>
      [SecurityCritical]
      internal static string GetRegularPathCore(string path, GetFullPathOptions options, bool allowEmpty)
      {
         if (null == path)
         {
            throw new ArgumentNullException("path");
         }

         if (!allowEmpty && Utils.IsNullOrWhiteSpace(path))
         {
            throw new ArgumentException(Resources.Path_Is_Zero_Length_Or_Only_White_Space, "path");
         }

         if (options != GetFullPathOptions.None)
         {
            path = ApplyFullPathOptions(path, options);
         }


         // 最も一般的なケース: パスが '\' で始まらなければプレフィックスなし → 即座に返す
         // 全プレフィックス（\\?\, \\.\, \??\）は '\' で始まる
         if (path.Length == 0 || path[0] != DirectorySeparatorChar)
            return path;

         // DosDeviceUncPrefix (\??\UNC\) — LongPathPrefix で始まらないため別途チェック
         if (path.StartsWith(DosDeviceUncPrefix, StringComparison.OrdinalIgnoreCase))
            return UncPrefix + path.Substring(DosDeviceUncPrefix.Length);

         // LogicalDrivePrefix (\\.\) — LongPathPrefix で始まらないため別途チェック
         if (path.StartsWith(LogicalDrivePrefix, StringComparison.Ordinal))
            return path.Substring(LogicalDrivePrefix.Length);

         // NonInterpretedPathPrefix (\??\) — LongPathPrefix で始まらないため別途チェック
         if (path.StartsWith(NonInterpretedPathPrefix, StringComparison.Ordinal))
            return path.Substring(NonInterpretedPathPrefix.Length);

         // LongPathPrefix (\\?\) で始まらないパスは即座に返す（UNCパス等）
         if (!path.StartsWith(LongPathPrefix, StringComparison.Ordinal))
            return path;

         // 以下、LongPathPrefix で始まるパスのみ到達
         if (path.StartsWith(LongPathUncPrefix, StringComparison.OrdinalIgnoreCase))
            return UncPrefix + path.Substring(LongPathUncPrefix.Length);

         if (path.StartsWith(GlobalRootPrefix, StringComparison.OrdinalIgnoreCase) ||
             path.StartsWith(VolumePrefix, StringComparison.OrdinalIgnoreCase))
            return path;

         return path.Substring(LongPathPrefix.Length);
      }
   }
}
