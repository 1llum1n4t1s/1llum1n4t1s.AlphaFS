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
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Volume
   {
      /// <summary>[AlphaFS] シリアル番号を比較して、2 つのファイルシステムオブジェクトのボリュームが同じかどうかを判断します。</summary>
      /// <param name="path1">フルパス情報を持つ最初のファイルシステムオブジェクト。</param>
      /// <param name="path2">フルパス情報を持つ 2 番目のファイルシステムオブジェクト。</param>
      /// <returns>両方のファイルシステムオブジェクトが同じボリュームにある場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      public static bool IsSameVolume(string path1, string path2)
      {
         try
         {
            var volInfo1 = new VolumeInfo(GetVolumePathName(path1), true, true);
            var volInfo2 = new VolumeInfo(GetVolumePathName(path2), true, true);

            return volInfo1.SerialNumber.Equals(volInfo2.SerialNumber) || volInfo1.Guid.Equals(volInfo2.Guid, StringComparison.OrdinalIgnoreCase);
         }
         catch (Exception ex)
         {
            // ネットワーク切断 / アクセス権なし / 不正なパス等の場合は false（判定不能 = 異ボリューム扱い）を返す。
            // ただし完全無音化すると `Directory.Move` の Copy+Delete fallback を意図せず誘発するため、
            // 切り分け可能な形で Trace に記録する。
            System.Diagnostics.Trace.TraceWarning(
               string.Format(System.Globalization.CultureInfo.InvariantCulture,
                  "Volume.IsSameVolume: 判定失敗 path1='{0}' path2='{1}': {2}",
                  path1, path2, ex.Message));
         }

         return false;
      }
   }
}
