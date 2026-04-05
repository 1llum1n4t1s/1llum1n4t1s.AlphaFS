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
   public static partial class Directory
   {
      #region .NET

      /// <summary>指定されたディレクトリが最後にアクセスされた日時を協定世界時（UTC）で設定します。</summary>
      /// <param name="path">The directory for which to set the access date and time information.</param>
      /// <param name="lastAccessTimeUtc">A <see cref="DateTime"/> 最終アクセス日時に設定する値を含む of <paramref name="path"/>. この値はUTC時刻で表されます。</param>
      [SecurityCritical]
      public static void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc)
      {
         File.SetFsoDateTimeCore(null, true, path, null, lastAccessTimeUtc, null, false, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary>[AlphaFS] 指定されたディレクトリが最後にアクセスされた日時を協定世界時（UTC）で設定します。</summary>
      /// <param name="path">The directory for which to set the access date and time information.</param>
      /// <param name="lastAccessTimeUtc">A <see cref="DateTime"/> 最終アクセス日時に設定する値を含む of <paramref name="path"/>. この値はUTC時刻で表されます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc, PathFormat pathFormat)
      {
         File.SetFsoDateTimeCore(null, true, path, null, lastAccessTimeUtc, null, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定されたディレクトリが最後にアクセスされた日時を協定世界時（UTC）で設定します。</summary>
      /// <param name="path">The directory for which to set the access date and time information.</param>
      /// <param name="lastAccessTimeUtc">A <see cref="DateTime"/> 最終アクセス日時に設定する値を含む of <paramref name="path"/>. この値はUTC時刻で表されます。</param>
      /// <param name="modifyReparsePoint">If <c>true</c>, the date and time information will apply to the reparse point (symlink or junction) and not the file or directory linked to. No effect if <paramref name="path"/> does not refer to a reparse point.</param>
      [SecurityCritical]
      public static void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc, bool modifyReparsePoint)
      {
         File.SetFsoDateTimeCore(null, true, path, null, lastAccessTimeUtc, null, modifyReparsePoint, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定されたディレクトリが最後にアクセスされた日時を協定世界時（UTC）で設定します。</summary>
      /// <param name="path">The directory for which to set the access date and time information.</param>
      /// <param name="lastAccessTimeUtc">A <see cref="DateTime"/> 最終アクセス日時に設定する値を含む of <paramref name="path"/>. この値はUTC時刻で表されます。</param>
      /// <param name="modifyReparsePoint">If <c>true</c>, the date and time information will apply to the reparse point (symlink or junction) and not the file or directory linked to. No effect if <paramref name="path"/> does not refer to a reparse point.</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc, bool modifyReparsePoint, PathFormat pathFormat)
      {
         File.SetFsoDateTimeCore(null, true, path, null, lastAccessTimeUtc, null, modifyReparsePoint, pathFormat);
      }
   }
}
