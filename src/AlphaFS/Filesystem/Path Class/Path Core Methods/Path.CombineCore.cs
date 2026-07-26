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
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>文字列の配列を1つのパスに結合します。</summary>
      /// <returns>結合されたパス。</returns>
      /// <remarks>
      ///   <para>パラメータに空白が含まれている場合は解析されません。</para>
      ///   <para>したがって、path2 に空白が含まれている場合（例: " c:\\ "）、</para>
      ///   <para>Combine メソッドは path2 のみを返す代わりに path2 を path1 に追加します。</para>
      /// </remarks>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <param name="checkInvalidPathChars"><c>true</c> の場合、<paramref name="paths"/> の無効なパス文字をチェックしません。</param>
      /// <param name="paths">パスの構成部分の配列。</param>
      [SecurityCritical]
      internal static string CombineCore(bool checkInvalidPathChars, params string[] paths)
      {
         if (null == paths)
         {
            throw new ArgumentNullException("paths");
         }

         var capacity = 0;
         var num = 0;

         for (int index = 0, l = paths.Length; index < l; ++index)
         {
            if (null == paths[index])
            {
               throw new ArgumentNullException("paths");
            }

            if (paths[index].Length != 0)
            {
               if (IsPathRooted(paths[index], checkInvalidPathChars))
               {
                  num = index;
                  capacity = paths[index].Length;
               }

               else
               {
                  capacity += paths[index].Length;
               }


               var ch = paths[index][paths[index].Length - 1];

               // ディレクトリ区切りのみを見る (ボリューム区切り ':' は含めない)。
               // .NET Framework は "C:" + "file" を "C:file" (ドライブ相対) にしていたが、
               // .NET Core 以降の System.IO.Path.Combine は "C:\file" を返す。ドロップイン代替として現行挙動に合わせる。
               if (!IsDVsc(ch, false))
               {
                  ++capacity;
               }
            }
         }


         var buffer = new StringBuilder(capacity);

         for (var index = num; index < paths.Length; ++index)
         {
            if (paths[index].Length != 0)
            {
               if (buffer.Length == 0)
               {
                  buffer.Append(paths[index]);
               }

               else
               {
                  var ch = buffer[buffer.Length - 1];

                  // capacity 計算と同じ判定にそろえる (ボリューム区切り ':' は区切りとみなさない)。
                  if (!IsDVsc(ch, false))
                  {
                     buffer.Append(DirectorySeparatorChar);
                  }

                  buffer.Append(paths[index]);
               }
            }
         }

         return buffer.ToString();
      }
   }
}
