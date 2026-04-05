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
   public static partial class Path
   {
      /// <summary>指定されたパス文字列のファイル名と拡張子を返します。</summary>
      /// <returns>
      ///   <paramref name="path"/> 内の最後のディレクトリ区切り文字の後の文字。<paramref name="path"/> の最後の文字が
      ///   ディレクトリ区切り文字またはボリューム区切り文字の場合、このメソッドは <c>string.Empty</c> を返します。path が null の場合、このメソッドは null を返します。
      /// </returns>
      /// <exception cref="ArgumentException"/>
      /// <param name="path">ファイル名と拡張子を取得するパス文字列。</param>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="path"/> の無効なパス文字をチェックします。</param>
      [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0", Justification = "Utils.IsNullOrWhiteSpace validates arguments.")]
      [SecurityCritical]
      internal static string GetFileNameCore(string path, bool checkInvalidPathChars)
      {
         if (null == path)
         {
            return null;
         }

         if (checkInvalidPathChars)
         {
            CheckInvalidPathChars(path, false, true);
         }

         var length = path.Length;
         var index = length;

         while (--index >= 0)
         {
            var ch = path[index];

            if (IsDVsc(ch, null))
            {
               return path.Substring(index + 1, length - index - 1);
            }
         }

         return path;
      }
   }
}
