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
   public static partial class File
   {
      /// <summary>[AlphaFS] 指定されたファイルのすべての日時スタンプを協定世界時(UTC)で一度に設定します。</summary>
      /// <param name="path">日時情報を設定するファイル。</param>
      /// <param name="creationTimeUtc"><paramref name="path"/>の作成日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastAccessTimeUtc"><paramref name="path"/>の最終アクセス日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastWriteTimeUtc"><paramref name="path"/>の最終書き込み日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SecurityCritical]
      public static void SetTimestampsUtc(string path, DateTime creationTimeUtc, DateTime lastAccessTimeUtc, DateTime lastWriteTimeUtc, PathFormat pathFormat)
      {
         SetFsoDateTimeCore(null, false, path, creationTimeUtc, lastAccessTimeUtc, lastWriteTimeUtc, false, pathFormat);
      }

      /// <summary>[AlphaFS] 指定されたファイルのすべての日時スタンプを協定世界時(UTC)で一度に設定します。</summary>
      /// <param name="path">日時情報を設定するファイル。</param>
      /// <param name="creationTimeUtc"><paramref name="path"/>の作成日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastAccessTimeUtc"><paramref name="path"/>の最終アクセス日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastWriteTimeUtc"><paramref name="path"/>の最終書き込み日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      [SecurityCritical]
      public static void SetTimestampsUtc(string path, DateTime creationTimeUtc, DateTime lastAccessTimeUtc, DateTime lastWriteTimeUtc)
      {
         SetFsoDateTimeCore(null, false, path, creationTimeUtc, lastAccessTimeUtc, lastWriteTimeUtc, false, PathFormat.RelativePath);
      }

      /// <summary>[AlphaFS] 指定されたファイルのすべての日時スタンプを協定世界時(UTC)で一度に設定します。</summary>
      /// <param name="path">日時情報を設定するファイル。</param>
      /// <param name="creationTimeUtc"><paramref name="path"/>の作成日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastAccessTimeUtc"><paramref name="path"/>の最終アクセス日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="lastWriteTimeUtc"><paramref name="path"/>の最終書き込み日時に設定する値を含む<see cref="DateTime"/>。この値はUTC時刻で表されます。</param>
      /// <param name="modifyReparsePoint"><c>true</c>の場合、日時情報はリパースポイント(シンボリックリンクまたはジャンクション)に適用され、リンク先のファイルまたはディレクトリには適用されません。<paramref name="path"/>がリパースポイントを参照していない場合は効果がありません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SecurityCritical]
      public static void SetTimestampsUtc(string path, DateTime creationTimeUtc, DateTime lastAccessTimeUtc, DateTime lastWriteTimeUtc, bool modifyReparsePoint, PathFormat pathFormat)
      {
         SetFsoDateTimeCore(null, false, path, creationTimeUtc, lastAccessTimeUtc, lastWriteTimeUtc, modifyReparsePoint, pathFormat);
      }
   }
}
