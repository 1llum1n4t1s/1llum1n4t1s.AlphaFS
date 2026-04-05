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
      /// <summary>ファイルの基本情報を格納します。ファイルハンドルに使用されます。</summary>
      /// <remarks>
      ///   <para><see cref="LastAccessTime"/>、<see cref="ChangeTime"/>、または <see cref="LastWriteTime"/> に -1 を指定すると、</para>
      ///   <para>現在のハンドルでの操作が指定されたフィールドに影響を与えないことを示します。</para>
      ///   <para>（つまり、<see cref="LastWriteTime"/> に -1 を指定すると、現在のハンドルで実行される書き込みによって
      ///   <see cref="LastWriteTime"/> は影響を受けません。）</para>
      /// </remarks>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct FILE_BASIC_INFO
      {
         /// <summary><see cref="FILETIME"/> 形式でのファイル作成日時。
         /// <para>1601年1月1日 (UTC) からの 100 ナノ秒間隔の数を表す 64 ビット値です。</para>
         /// </summary>
         public FILETIME CreationTime;

         /// <summary><see cref="FILETIME"/> 形式でのファイル最終アクセス日時。</summary>
         public FILETIME LastAccessTime;

         /// <summary><see cref="FILETIME"/> 形式でのファイル最終書き込み日時。</summary>
         public FILETIME LastWriteTime;

         /// <summary><see cref="FILETIME"/> 形式でのファイル変更日時。</summary>
         public FILETIME ChangeTime;

         /// <summary>ファイル属性。</summary>
         /// <remarks>SetFileInformationByHandle に渡される <see cref="FILE_BASIC_INFO"/> 構造体でこの値が 0 に設定されている場合、属性は変更されません。</remarks>
         public FileAttributes FileAttributes;
      }
   }
}