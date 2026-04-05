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
      /// <summary>[AlphaFS] 指定されたパスの一意のボリューム名を取得します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="volumePathName">
      ///   パス文字列。".." などの絶対および相対ファイル名とディレクトリ名の両方がこのパスで使用できます。
      ///   ボリューム修飾子なしの相対ファイル名またはディレクトリ名を指定した場合、GetUniqueVolumeNameForPath は現在のボリュームの
      ///   ドライブ文字を返します。
      /// </param>
      /// <returns>
      ///   <para>"\\?\Volume{GUID}\" 形式の一意のボリューム名を返します。</para>
      ///   <para>エラーの場合または利用できない場合は <c>null</c>。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      public static string GetUniqueVolumeNameForPath(string volumePathName)
      {
         if (Utils.IsNullOrWhiteSpace(volumePathName))
         {
            throw new ArgumentNullException("volumePathName");
         }

         try
         {
            return GetVolumeGuid(GetVolumePathName(volumePathName));
         }
         catch
         {
            return null;
         }
      }
   }
}
