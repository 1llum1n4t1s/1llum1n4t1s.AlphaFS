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
using System.IO;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>[AlphaFS] 指定されたディレクトリ内のファイルに関する情報を格納します。ディレクトリハンドルに使用されます。</summary>
   [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dir")]
   [Serializable]
   [SecurityCritical]
   public sealed class FileIdBothDirectoryInfo
   {
      internal FileIdBothDirectoryInfo(NativeMethods.FILE_ID_BOTH_DIR_INFO fibdi, string fileName)
      {
         CreationTimeUtc = DateTime.FromFileTimeUtc(fibdi.CreationTime);
         LastAccessTimeUtc = DateTime.FromFileTimeUtc(fibdi.LastAccessTime);
         LastWriteTimeUtc = DateTime.FromFileTimeUtc(fibdi.LastWriteTime);
         ChangeTimeUtc = DateTime.FromFileTimeUtc(fibdi.ChangeTime);

         AllocationSize = fibdi.AllocationSize;
         EndOfFile = fibdi.EndOfFile;
         ExtendedAttributesSize = fibdi.EaSize;
         
         FileAttributes = fibdi.FileAttributes;
         FileId = fibdi.FileId;
         FileIndex = fibdi.FileIndex;
         FileName = fileName;

         // ShortNameLengthは短い名前のバイト数です。Unicode文字列なので2で除算する必要があります。
         ShortName = new string(fibdi.ShortName, 0, fibdi.ShortNameLength / UnicodeEncoding.CharSize);
      }




      /// <summary>ファイルに割り当てられたバイト数。この値は通常、基盤となる物理デバイスのセクターまたはクラスターサイズの倍数です。</summary>
      public long AllocationSize { get; set; }


      /// <summary>このエントリが変更された時刻を取得します。</summary>
      /// <value>このエントリが変更された時刻。</value>
      public DateTime ChangeTime
      {
         get { return ChangeTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが変更された協定世界時（UTC）での時刻を取得します。</summary>
      /// <value>このエントリが変更されたUTC時刻。</value>
      public DateTime ChangeTimeUtc { get; set; }


      /// <summary>このエントリが作成された時刻を取得します。</summary>
      /// <value>このエントリが作成された時刻。</value>
      public DateTime CreationTime
      {
         get { return CreationTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが作成された協定世界時（UTC）での時刻を取得します。</summary>
      /// <value>このエントリが作成されたUTC時刻。</value>
      public DateTime CreationTimeUtc { get; set; }


      /// <summary>ファイルの拡張属性のサイズ。</summary>
      public int ExtendedAttributesSize { get; set; }


      /// <summary>ファイルの先頭から末尾までのバイトオフセットとして表される、ファイル終端の絶対位置。
      /// この値はゼロベースであるため、実際にはファイル内の最初の空きバイトを指します。つまり、<b>EndOfFile</b>は
      /// ファイル内の最後の有効なバイトの直後のバイトへのオフセットです。
      /// </summary>
      public long EndOfFile { get; set; }


      /// <summary>ファイル属性。</summary>
      public FileAttributes FileAttributes { get; set; }


      /// <summary>ファイルID。</summary>
      public long FileId { get; set; }


      /// <summary>親ディレクトリ内でのファイルのバイトオフセット。NTFSなど、親ディレクトリ内でのファイルの位置が
      /// 固定されておらず、ソート順序を維持するためにいつでも変更される可能性があるファイルシステムでは、このメンバーは未定義です。
      /// </summary>
      public long FileIndex { get; set; }


      /// <summary>ファイル名。</summary>
      public string FileName { get; set; }


      /// <summary>このエントリに最後にアクセスした時刻を取得します。</summary>
      /// <value>このエントリに最後にアクセスした時刻。</value>
      public DateTime LastAccessTime
      {
         get { return LastAccessTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリに最後にアクセスした協定世界時（UTC）での時刻を取得します。</summary>
      /// <value>このエントリに最後にアクセスしたUTC時刻。</value>
      public DateTime LastAccessTimeUtc { get; set; }


      /// <summary>このエントリが最後に変更された時刻を取得します。</summary>
      /// <value>このエントリが最後に変更された時刻。</value>
      public DateTime LastWriteTime
      {
         get { return LastWriteTimeUtc.ToLocalTime(); }
      }


      /// <summary>このエントリが最後に変更された協定世界時（UTC）での時刻を取得します。</summary>
      /// <value>このエントリが最後に変更されたUTC時刻。</value>
      public DateTime LastWriteTimeUtc { get; set; }


      /// <summary>ファイルの8.3形式の短い名前（例: FILENAME.TXT）。</summary>
      public string ShortName { get; set; }
   }
}
