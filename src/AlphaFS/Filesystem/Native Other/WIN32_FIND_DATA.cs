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
using System.IO;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>FindFirstFile、FindFirstFileEx、または FindNextFile 関数によって検出されたファイルに関する情報を格納します。</summary>
      /// <remarks>
      /// ファイルに長いファイル名がある場合、完全な名前は cFileName メンバーに表示され、8.3 形式の切り詰められた名前は
      /// cAlternateFileName メンバーに表示されます。それ以外の場合、cAlternateFileName は空です。FindFirstFileEx 関数が fInfoLevelId パラメーターに
      /// FindExInfoBasic の値で呼び出された場合、cAlternateFileName メンバーは常に <c>null</c> 文字列値を格納します。これは FindNextFile 関数への
      /// すべての後続の呼び出しでも同様です。ファイル名の 8.3 形式バージョンを取得する代替方法として、GetShortPathName 関数を使用できます。
      /// ファイル名の詳細については、File Names, Paths, and Namespaces を参照してください。
      /// </remarks>
      /// <remarks>
      /// すべてのファイルシステムが作成日時と最終アクセス日時を記録できるわけではなく、同じ方法で記録するわけでもありません。
      /// 例えば、FAT ファイルシステムでは、作成日時の解像度は 10 ミリ秒、書き込み日時の解像度は 2 秒、
      /// アクセス日時の解像度は 1 日です。NTFS ファイルシステムでは、最終アクセス日時の更新を最後のアクセスから最大 1 時間遅延させます。
      /// 詳細については、File Times を参照してください。
      /// </remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      [Serializable]
      internal struct WIN32_FIND_DATA
      {
         /// <summary>ファイルのファイル属性。</summary>
         public FileAttributes dwFileAttributes;


         /// <summary>ファイルまたはディレクトリが作成された日時を指定する <see cref="FILETIME"/> 構造体。
         /// 基になるファイルシステムが作成日時をサポートしない場合、このメンバーはゼロです。</summary>
         public FILETIME ftCreationTime;


         /// <summary><see cref="FILETIME"/> 構造体。
         /// ファイルの場合、最後にファイルが読み取り、書き込み、または実行可能ファイルでは実行された日時を指定します。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。基になるファイルシステムが最終アクセス日時をサポートしない場合、このメンバーはゼロです。
         /// FAT ファイルシステムでは、ファイルとディレクトリの両方で指定された日付は正しいですが、時刻は常に午前0時に設定されます。
         /// </summary>
         public FILETIME ftLastAccessTime;


         /// <summary><see cref="FILETIME"/> 構造体。
         /// ファイルの場合、最後にファイルが書き込み、切り詰め、または上書きされた日時を指定します（例: WriteFile や SetEndOfFile が使用された場合）。
         /// ファイル属性やセキュリティ記述子が変更された場合、日付と時刻は更新されません。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。基になるファイルシステムが最終書き込み日時をサポートしない場合、このメンバーはゼロです。
         /// </summary>
         public FILETIME ftLastWriteTime;


         /// <summary>ファイルサイズの上位 DWORD。このメンバーはディレクトリには意味がありません。
         /// ファイルサイズが MAXDWORD より大きい場合を除き、この値はゼロです。
         /// ファイルのサイズは (nFileSizeHigh * (MAXDWORD+1)) + nFileSizeLow に等しくなります。
         /// </summary>
         public uint nFileSizeHigh;


         /// <summary>ファイルサイズの下位 DWORD。このメンバーはディレクトリには意味がありません。</summary>
         public uint nFileSizeLow;


         /// <summary>dwFileAttributes メンバーに FILE_ATTRIBUTE_REPARSE_POINT 属性が含まれている場合、このメンバーはリパースポイントタグを指定します。
         /// それ以外の場合、この値は未定義であり使用すべきではありません。
         /// </summary>
         public readonly ReparsePointTag dwReserved0;


         /// <summary>将来の使用のために予約されています。</summary>
         private readonly uint dwReserved1;


         /// <summary>ファイルの名前。</summary>
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MaxPath)] public string cFileName;


         /// <summary>ファイルの代替名。この名前は従来の 8.3 ファイル名形式です。</summary>
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public readonly string cAlternateFileName;
      }
   }
}
