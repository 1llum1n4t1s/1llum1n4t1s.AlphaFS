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

using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>
      ///   指定されたディスクに関する情報（ディスク上の空き容量を含む）を取得します。
      /// </summary>
      /// <remarks>
      ///   <para>シンボリックリンクの動作: パスがシンボリックリンクを指している場合、操作はターゲットに対して実行されます。</para>
      ///   <para>このパラメータが UNC 名の場合、末尾にバックスラッシュを含める必要があります（例: "\\MyServer\MyShare\"）。</para>
      ///   <para>さらに、ドライブ指定には末尾にバックスラッシュが必要です（例: "C:\"）。</para>
      ///   <para>呼び出し元のアプリケーションは、このディレクトリに対する FILE_LIST_DIRECTORY アクセス権を持っている必要があります。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpRootPathName">ルートファイルの完全パス名。</param>
      /// <param name="lpSectorsPerCluster">[out] クラスタあたりのセクタ数。</param>
      /// <param name="lpBytesPerSector">[out] セクタあたりのバイト数。</param>
      /// <param name="lpNumberOfFreeClusters">[out] 空きクラスタ数。</param>
      /// <param name="lpTotalNumberOfClusters">[out] クラスタの合計数。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetDiskFreeSpaceW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetDiskFreeSpace([MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName, [MarshalAs(UnmanagedType.U4)] out int lpSectorsPerCluster, [MarshalAs(UnmanagedType.U4)] out int lpBytesPerSector, [MarshalAs(UnmanagedType.U4)] out int lpNumberOfFreeClusters, [MarshalAs(UnmanagedType.U4)] out uint lpTotalNumberOfClusters);

      /// <summary>
      ///   ディスクボリュームで利用可能な容量に関する情報を取得します。合計容量、
      ///   <para>合計空き容量、および呼び出しスレッドに関連付けられたユーザーが利用可能な合計空き容量が含まれます。</para>
      /// </summary>
      /// <remarks>
      ///   <para>シンボリックリンクの動作: パスがシンボリックリンクを指している場合、操作はターゲットに対して実行されます。</para>
      ///   <para>GetDiskFreeSpaceEx 関数は、ディスクが CD-RW ドライブの未書き込み CD でない限り、すべての CD リクエストに対して
      ///   lpTotalNumberOfFreeBytes および lpFreeBytesAvailable にゼロ (0) を返します。</para>
      ///   <para>このパラメータが UNC 名の場合、末尾にバックスラッシュを含める必要があります（例: "\\MyServer\MyShare\"）。</para>
      ///   <para>このパラメータはディスク上のルートディレクトリを指定する必要はありません。</para>
      ///   <para>この関数はディスク上の任意のディレクトリを受け入れます。</para>
      ///   <para>呼び出し元のアプリケーションは、このディレクトリに対する FILE_LIST_DIRECTORY アクセス権を持っている必要があります。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      /// <param name="lpDirectoryName">ディレクトリのパス名。</param>
      /// <param name="lpFreeBytesAvailable">[out] 利用可能な空きバイト数。</param>
      /// <param name="lpTotalNumberOfBytes">[out] 合計バイト数。</param>
      /// <param name="lpTotalNumberOfFreeBytes">[out] 合計空きバイト数。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロ (0) です。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetDiskFreeSpaceExW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool GetDiskFreeSpaceEx([MarshalAs(UnmanagedType.LPWStr)] string lpDirectoryName, [MarshalAs(UnmanagedType.U8)] out long lpFreeBytesAvailable, [MarshalAs(UnmanagedType.U8)] out long lpTotalNumberOfBytes, [MarshalAs(UnmanagedType.U8)] out long lpTotalNumberOfFreeBytes);
   }
}
