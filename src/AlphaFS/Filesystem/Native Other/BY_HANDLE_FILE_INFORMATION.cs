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
      /// <summary>GetFileInformationByHandle 関数が取得する情報を格納します。</summary>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct BY_HANDLE_FILE_INFORMATION
      {
         /// <summary>ファイル属性。</summary>
         public readonly FileAttributes dwFileAttributes;


         /// <summary>ファイルまたはディレクトリが作成された日時を指定する <see cref="FILETIME"/> 構造体。</summary>
         public readonly FILETIME ftCreationTime;


         /// <summary><see cref="FILETIME"/> 構造体。ファイルの場合、最後にファイルが読み取りまたは書き込みされた日時を指定します。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。
         /// ファイルとディレクトリの両方において、指定された日付は正しいですが、時刻は常に午前0時に設定されます。
         /// </summary>
         public readonly FILETIME ftLastAccessTime;


         /// <summary><see cref="FILETIME"/> 構造体。ファイルの場合、最後にファイルが書き込まれた日時を指定します。
         /// ディレクトリの場合、ディレクトリが作成された日時を指定します。</summary>
         public readonly FILETIME ftLastWriteTime;


         /// <summary>ファイルを含むボリュームのシリアル番号。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint dwVolumeSerialNumber;


         /// <summary>ファイルサイズの上位部分。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint nFileSizeHigh;


         /// <summary>ファイルサイズの下位部分。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint nFileSizeLow;

         /// <summary>このファイルへのリンク数。FAT ファイルシステムではこのメンバーは常に 1 です。NTFS ファイルシステムでは 1 より大きくなる場合があります。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint nNumberOfLinks;

         /// <summary>ファイルに関連付けられた一意の識別子の上位部分。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint nFileIndexHigh;

         /// <summary>ファイルに関連付けられた一意の識別子の下位部分。</summary>
         [MarshalAs(UnmanagedType.U4)] public readonly uint nFileIndexLow;
      }
   }
}
