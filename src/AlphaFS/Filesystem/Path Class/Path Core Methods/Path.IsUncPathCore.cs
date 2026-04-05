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
      /// <summary>パス文字列が有効なUNC（汎用名前付け規則）パスであるかどうかを判定します。オプションで無効なパス文字のチェックをスキップできます。</summary>
      /// <returns>指定されたパスがUNC（汎用名前付け規則）パスの場合は <c>true</c>、それ以外は <c>false</c>。</returns>
      /// <param name="path">チェックするパス。</param>
      /// <param name="isRegularPath"><c>true</c> の場合、<paramref name="path"/> が既に通常のパス形式であることを示します。</param>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="path"/> の無効なパス文字をチェックします。</param>
      [SecurityCritical]
      internal static bool IsUncPathCore(string path, bool isRegularPath, bool checkInvalidPathChars)
      {
         if (!isRegularPath)
         {
            path = GetRegularPathCore(path, checkInvalidPathChars ? GetFullPathOptions.CheckInvalidPathChars : GetFullPathOptions.None, false);
         }

         else if (checkInvalidPathChars)
         {
            CheckInvalidPathChars(path, false, false);
         }


         return Uri.TryCreate(path, UriKind.Absolute, out var uri) && uri.IsUnc;
      }
   }
}
