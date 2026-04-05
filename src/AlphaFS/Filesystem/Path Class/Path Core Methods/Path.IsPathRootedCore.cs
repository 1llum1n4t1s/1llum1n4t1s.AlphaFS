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
      /// <summary>指定されたパス文字列に絶対パス情報と相対パス情報のどちらが含まれているかを示す値を取得します。</summary>
      /// <returns><paramref name="path"/> にルートが含まれている場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <remarks>
      ///   IsPathRooted メソッドは、最初の文字が <see cref="DirectorySeparatorChar"/> のようなディレクトリ区切り文字の場合、
      ///   またはパスがドライブ文字とコロン（<see cref="VolumeSeparatorChar"/>）で始まる場合に <c>true</c> を返します。
      ///   例えば、"\\MyDir\\MyFile.txt"、"C:\\MyDir"、"C:MyDir" のようなパス文字列に対して <c>true</c> を返します。
      ///   "MyDir" のようなパス文字列に対しては <c>false</c> を返します。
      /// </remarks>
      /// <remarks>このメソッドはパスやファイル名が存在するかどうかを検証しません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">テストするパス。パスには <see cref="GetInvalidPathChars"/> で定義されている文字を含めることはできません。</param>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="path"/> の無効なパス文字をチェックします。</param>
      [SecurityCritical]
      internal static bool IsPathRootedCore(string path, bool checkInvalidPathChars)
      {
         if (null != path)
         {
            if (checkInvalidPathChars)
            {
               CheckInvalidPathChars(path, false, true);
            }

            var length = path.Length;

            if (length >= 1 && IsDVsc(path[0], false) || length >= 2 && IsDVsc(path[1], true))
            {
               return true;
            }
         }

         return false;
      }
   }
}
