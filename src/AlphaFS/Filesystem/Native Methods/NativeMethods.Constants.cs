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
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      public static readonly bool IsAtLeastWindows8 = OperatingSystem.IsAtLeast(OperatingSystem.EnumOsName.Windows8);
      public static readonly bool IsAtLeastWindows7 = OperatingSystem.IsAtLeast(OperatingSystem.EnumOsName.Windows7);
      public static readonly bool IsAtLeastWindowsVista = OperatingSystem.IsAtLeast(OperatingSystem.EnumOsName.WindowsVista);

      /// <summary>FindFirstFileEx 関数がショートファイル名を照会せず、列挙速度全体を向上させます。
      /// <para>&#160;</para>
      /// <remarks>
      /// <para>データは <see cref="WIN32_FIND_DATA"/> 構造体で返され、</para>
      /// <para>cAlternateFileName メンバーは常に NULL 文字列です。</para>
      /// <para>この値は Windows Server 2008 R2 および Windows 7 以降でサポートされます。</para>
      /// </remarks>
      /// </summary>
      public static readonly FINDEX_INFO_LEVELS FindexInfoLevel = IsAtLeastWindows7 ? FINDEX_INFO_LEVELS.Basic : FINDEX_INFO_LEVELS.Standard;

      /// <summary>ディレクトリクエリにより大きなバッファを使用し、検索操作のパフォーマンスを向上させます。</summary>
      /// <remarks>この値は Windows Server 2008 R2 および Windows 7 以降でサポートされます。</remarks>
      public static readonly FIND_FIRST_EX_FLAGS UseLargeCache = IsAtLeastWindows7 ? FIND_FIRST_EX_FLAGS.LARGE_FETCH : FIND_FIRST_EX_FLAGS.NONE;

      /// <summary>DefaultFileBufferSize = 65536; ファイルの読み書きに使用されるデフォルトのバッファサイズ。</summary>
      public const int DefaultFileBufferSize = 65536;

      /// <summary>DefaultNativeQueryBufferSize = 4096; ネイティブAPIクエリ（デバイス情報、ボリューム、シェル等）用のスクラッチバッファサイズ。
      /// ファイルI/Oには<see cref="DefaultFileBufferSize"/>を使用すること。</summary>
      internal const int DefaultNativeQueryBufferSize = 4096;

      /// <summary>ディレクトリ列挙キューの初期容量。</summary>
      internal const int DefaultDirectoryQueueCapacity = 64;

      /// <summary>DefaultFileEncoding = Encoding.UTF8; ファイルの読み書きに使用されるデフォルトのエンコーディング。</summary>
      public static readonly Encoding DefaultFileEncoding = Encoding.UTF8;

      /// <summary>MaxDirectoryLength = 255</summary>
      internal const int MaxDirectoryLength = 255;

      /// <summary>MaxPath = 260
      /// 指定されたパス、ファイル名、またはその両方がシステム定義の最大長を超えています。
      /// 例えば、Windows ベースのプラットフォームでは、パスは 248 文字未満、ファイル名は 260 文字未満でなければなりません。
      /// </summary>
      internal const int MaxPath = 260;

      /// <summary>MaxPathUnicode = 32700</summary>
      internal const int MaxPathUnicode = 32700;


      /// <summary>例外発生時に "System.OverflowException: 算術演算でオーバーフローが発生しました。" を防ぐためにビットシフトが必要です。</summary>
      internal const int OverflowExceptionBitShift = 65535;


      /// <summary>無効な FileAttributes = -1</summary>
      internal const FileAttributes InvalidFileAttributes = (FileAttributes) (-1);




      /// <summary>MAXIMUM_REPARSE_DATA_BUFFER_SIZE = 16384</summary>
      internal const int MAXIMUM_REPARSE_DATA_BUFFER_SIZE = 16384;

      /// <summary>REPARSE_DATA_BUFFER_HEADER_SIZE = 8</summary>
      internal const int REPARSE_DATA_BUFFER_HEADER_SIZE = 8;


      private const int DeviceIoControlMethodBuffered = 0;
      private const int DeviceIoControlFileDeviceFileSystem = 9;

      // <summary>ファイルごと・ディレクトリごとの圧縮をサポートするファイルシステムのボリューム上で、ファイルまたはディレクトリの圧縮状態を設定するコマンド。</summary>
      internal const int FSCTL_SET_COMPRESSION = (DeviceIoControlFileDeviceFileSystem << 16) | (16 << 2) | DeviceIoControlMethodBuffered | (int) (FileAccess.Read | FileAccess.Write) << 14;

      // <summary>リパースポイントデータブロックを設定するコマンド。</summary>
      internal const int FSCTL_SET_REPARSE_POINT = (DeviceIoControlFileDeviceFileSystem << 16) | (41 << 2) | DeviceIoControlMethodBuffered | (0 << 14);
      
      /// <summary>リパースポイントデータベースを削除するコマンド。</summary>
      internal const int FSCTL_DELETE_REPARSE_POINT = (DeviceIoControlFileDeviceFileSystem << 16) | (43 << 2) | DeviceIoControlMethodBuffered | (0 << 14);

      /// <summary>リパースポイントデータブロックを取得するコマンド。</summary>
      internal const int FSCTL_GET_REPARSE_POINT = (DeviceIoControlFileDeviceFileSystem << 16) | (42 << 2) | DeviceIoControlMethodBuffered | (0 << 14);
   }
}
