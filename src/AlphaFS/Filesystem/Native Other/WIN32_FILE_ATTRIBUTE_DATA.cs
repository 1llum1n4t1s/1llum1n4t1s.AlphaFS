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

using System.IO;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>WIN32_FILE_ATTRIBUTE_DATA 構造体は、ファイルまたはディレクトリの属性情報を格納します。GetFileAttributesEx 関数がこの構造体を使用します。</summary>
      /// <remarks>
      /// すべてのファイルシステムが作成日時と最終アクセス日時を記録できるわけではなく、同じ方法で記録するわけでもありません。
      /// 例えば、FAT ファイルシステムでは、作成日時の解像度は 10 ミリ秒、書き込み日時の解像度は 2 秒、
      /// アクセス日時の解像度は 1 日です。NTFS ファイルシステムでは、アクセス日時の解像度は 1 時間です。
      /// 詳細については、File Times を参照してください。
      /// </remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct WIN32_FILE_ATTRIBUTE_DATA
      {
         public WIN32_FILE_ATTRIBUTE_DATA(WIN32_FIND_DATA findData)
         {
            dwFileAttributes = findData.dwFileAttributes;
            ftCreationTime = findData.ftCreationTime;
            ftLastAccessTime = findData.ftLastAccessTime;
            ftLastWriteTime = findData.ftLastWriteTime;
            nFileSizeHigh = findData.nFileSizeHigh;
            nFileSizeLow = findData.nFileSizeLow;
         }

         /// <summary>ファイルのファイル属性。</summary>
         [MarshalAs(UnmanagedType.I4)] public FileAttributes dwFileAttributes;

         /// <summary>ファイルまたはディレクトリが作成された日時を指定する <see cref="FILETIME"/> 構造体。
         /// 基になるファイルシステムが作成日時をサポートしない場合、このメンバーはゼロです。</summary>
         public readonly FILETIME ftCreationTime;

         /// <summary><see cref="FILETIME"/> 構造体。
         /// ファイルの場合、最後にファイルが読み取り、書き込み、または実行可能ファイルでは実行された日時を指定します。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。基になるファイルシステムが最終アクセス日時をサポートしない場合、このメンバーはゼロです。
         /// FAT ファイルシステムでは、ファイルとディレクトリの両方で指定された日付は正しいですが、時刻は常に午前0時に設定されます。
         /// </summary>
         public readonly FILETIME ftLastAccessTime;

         /// <summary><see cref="FILETIME"/> 構造体。
         /// ファイルの場合、最後にファイルが書き込み、切り詰め、または上書きされた日時を指定します（例: WriteFile や SetEndOfFile が使用された場合）。
         /// ファイル属性やセキュリティ記述子が変更された場合、日付と時刻は更新されません。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。基になるファイルシステムが最終書き込み日時をサポートしない場合、このメンバーはゼロです。
         /// </summary>
         public readonly FILETIME ftLastWriteTime;

         /// <summary>ファイルサイズの上位 DWORD。このメンバーはディレクトリには意味がありません。
         /// ファイルサイズが MAXDWORD より大きい場合を除き、この値はゼロです。
         /// ファイルのサイズは (nFileSizeHigh * (MAXDWORD+1)) + nFileSizeLow に等しくなります。
         /// </summary>
         public readonly uint nFileSizeHigh;

         /// <summary>ファイルサイズの下位 DWORD。このメンバーはディレクトリには意味がありません。</summary>
         public readonly uint nFileSizeLow;

         /// <summary>ファイルサイズ。</summary>
         public long FileSize
         {
            get { return ToLong(nFileSizeHigh, nFileSizeLow); }
         }
      }
   }
}
