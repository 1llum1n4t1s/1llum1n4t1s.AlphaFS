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
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>
      ///   新しいディレクトリを作成します。
      ///   <para>基盤となるファイルシステムがファイルおよびディレクトリのセキュリティをサポートしている場合、</para>
      ///   <para>この関数は指定されたセキュリティ記述子を新しいディレクトリに適用します。</para>
      /// </summary>
      /// <remarks>
      ///   <para>NTFS ファイルシステムなど、一部のファイルシステムは個々のファイルおよびディレクトリの圧縮または暗号化をサポートします。</para>
      ///   <para>そのようなファイルシステムでフォーマットされたボリュームでは、新しいディレクトリは親ディレクトリの圧縮および暗号化属性を継承します。</para>
      ///   <para>アプリケーションは FILE_FLAG_BACKUP_SEMANTICS フラグを設定して <see cref="CreateFile"/> を呼び出すことで、ディレクトリのハンドルを取得できます。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      /// <param name="lpPathName">ファイルの完全パス名。</param>
      /// <param name="lpSecurityAttributes">セキュリティ属性。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateDirectoryW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CreateDirectory([MarshalAs(UnmanagedType.LPWStr)] string lpPathName, [MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpSecurityAttributes);

      /// <summary>
      ///   指定されたテンプレートディレクトリの属性を持つ新しいディレクトリを作成します。
      ///   <para>基盤となるファイルシステムがファイルおよびディレクトリのセキュリティをサポートしている場合、</para>
      ///   <para>この関数は指定されたセキュリティ記述子を新しいディレクトリに適用します。</para>
      ///   <para>新しいディレクトリは、指定されたテンプレートディレクトリのその他の属性を保持します。</para>
      /// </summary>
      /// <remarks>
      ///   <para>CreateDirectoryEx 関数を使用すると、他のディレクトリからストリーム情報を継承するディレクトリを作成できます。</para>
      ///   <para>この関数は、例えば、ディレクトリの内容を属性として適切に識別するために必要なリソースストリームを持つ</para>
      ///   <para>Macintosh ディレクトリを使用している場合に便利です。</para>
      ///   <para>NTFS ファイルシステムなど、一部のファイルシステムは個々のファイルおよびディレクトリの圧縮または暗号化をサポートします。</para>
      ///   <para>そのようなファイルシステムでフォーマットされたボリュームでは、新しいディレクトリは親ディレクトリの圧縮および暗号化属性を継承します。</para>
      ///   <para>FILE_FLAG_BACKUP_SEMANTICS フラグを設定して <see cref="CreateFile"/> 関数を呼び出すことで、ディレクトリのハンドルを取得できます。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpTemplateDirectory">テンプレートディレクトリのパス名。</param>
      /// <param name="lpPathName">ファイルの完全パス名。</param>
      /// <param name="lpSecurityAttributes">セキュリティ属性。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロ (0) です。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateDirectoryExW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CreateDirectoryEx([MarshalAs(UnmanagedType.LPWStr)] string lpTemplateDirectory, [MarshalAs(UnmanagedType.LPWStr)] string lpPathName, [MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpSecurityAttributes);

      /// <summary>
      ///   トランザクション操作として、指定されたテンプレートディレクトリの属性を持つ新しいディレクトリを作成します。
      ///   <para>基盤となるファイルシステムがファイルおよびディレクトリのセキュリティをサポートしている場合、</para>
      ///   <para>この関数は指定されたセキュリティ記述子を新しいディレクトリに適用します。</para>
      ///   <para>新しいディレクトリは、指定されたテンプレートディレクトリのその他の属性を保持します。</para>
      /// </summary>
      /// <remarks>
      ///   <para>CreateDirectoryTransacted 関数を使用すると、他のディレクトリからストリーム情報を継承するディレクトリを作成できます。</para>
      ///   <para>この関数は、例えば、ディレクトリの内容を属性として適切に識別するために必要なリソースストリームを持つ</para>
      ///   <para>Macintosh ディレクトリを使用している場合に便利です。</para>
      ///   <para>NTFS ファイルシステムなど、一部のファイルシステムは個々のファイルおよびディレクトリの圧縮または暗号化をサポートします。</para>
      ///   <para>そのようなファイルシステムでフォーマットされたボリュームでは、新しいディレクトリは親ディレクトリの圧縮および暗号化属性を継承します。</para>
      ///   <para>FILE_FLAG_BACKUP_SEMANTICS フラグを設定して <see cref="CreateFileTransacted"/> 関数を呼び出すことで、ディレクトリのハンドルを取得できます。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpTemplateDirectory">テンプレートディレクトリのパス名。</param>
      /// <param name="lpNewDirectory">新しいディレクトリのパス名。</param>
      /// <param name="lpSecurityAttributes">セキュリティ属性。</param>
      /// <param name="hTransaction">トランザクション。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロ (0) です。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      ///   <para>暗号化が無効な親ディレクトリで子ディレクトリを作成しようとすると、</para>
      ///   <para>この関数は ERROR_EFS_NOT_ALLOWED_IN_TRANSACTION で失敗します。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "CreateDirectoryTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool CreateDirectoryTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpTemplateDirectory, [MarshalAs(UnmanagedType.LPWStr)] string lpNewDirectory, [MarshalAs(UnmanagedType.LPStruct)] Security.NativeMethods.SecurityAttributes lpSecurityAttributes, SafeHandle hTransaction);

      /// <summary>
      ///   現在のプロセスのカレントディレクトリを取得します。
      /// </summary>
      /// <remarks>
      ///   <para>RemoveDirectory 関数はディレクトリを閉じる際に削除対象としてマークします。</para>
      ///   <para>そのため、ディレクトリへの最後のハンドルが閉じられるまでディレクトリは削除されません。</para>
      ///   <para>RemoveDirectory は、ターゲットの内容が空でなくてもディレクトリジャンクションを削除します。</para>
      ///   <para>この関数はターゲットオブジェクトの状態に関係なくディレクトリジャンクションを削除します。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      /// <param name="nBufferLength">カレントディレクトリ文字列のバッファの長さ (TCHAR 単位)。バッファの長さには終端のヌル文字のための領域を含める必要があります。</param>
      /// <param name="lpBuffer">
      ///   <para>カレントディレクトリ文字列を受け取るバッファへのポインタ。このヌル終端文字列はカレントディレクトリへの絶対パスを指定します。</para>
      ///   <para>必要なバッファサイズを判断するには、このパラメータを NULL に設定し、nBufferLength パラメータを 0 に設定します。</para>
      /// </param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値は終端のヌル文字を含まない、バッファに書き込まれた文字数を指定します。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Usage", "CA2205:UseManagedEquivalentsOfWin32Api")]
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "GetCurrentDirectoryW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint GetCurrentDirectory([MarshalAs(UnmanagedType.U4)] uint nBufferLength, StringBuilder lpBuffer);
      
      /// <summary>
      ///   既存の空のディレクトリを削除します。
      /// </summary>
      /// <remarks>
      ///   <para>RemoveDirectory 関数はディレクトリを閉じる際に削除対象としてマークします。</para>
      ///   <para>そのため、ディレクトリへの最後のハンドルが閉じられるまでディレクトリは削除されません。</para>
      ///   <para>RemoveDirectory は、ターゲットの内容が空でなくてもディレクトリジャンクションを削除します。</para>
      ///   <para>この関数はターゲットオブジェクトの状態に関係なくディレクトリジャンクションを削除します。</para>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      /// <param name="lpPathName">ファイルの完全パス名。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "RemoveDirectoryW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool RemoveDirectory([MarshalAs(UnmanagedType.LPWStr)] string lpPathName);

      /// <summary>
      ///   トランザクション操作として、既存の空のディレクトリを削除します。
      /// </summary>
      /// <remarks>
      ///   <para>RemoveDirectoryTransacted 関数はディレクトリを閉じる際に削除対象としてマークします。</para>
      ///   <para>そのため、ディレクトリへの最後のハンドルが閉じられるまでディレクトリは削除されません。</para>
      ///   <para>RemoveDirectory は、ターゲットの内容が空でなくてもディレクトリジャンクションを削除します。</para>
      ///   <para>この関数はターゲットオブジェクトの状態に関係なくディレクトリジャンクションを削除します。</para>
      ///   <para>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</para>
      /// </remarks>
      /// <param name="lpPathName">ファイルの完全パス名。</param>
      /// <param name="hTransaction">トランザクション。</param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "RemoveDirectoryTransactedW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool RemoveDirectoryTransacted([MarshalAs(UnmanagedType.LPWStr)] string lpPathName, SafeHandle hTransaction);

      /// <summary>
      ///   現在のプロセスのカレントディレクトリを変更します。
      /// </summary>
      /// <param name="lpPathName">
      ///   <para>新しいカレントディレクトリへのパス。このパラメータは相対パスまたは完全パスを指定できます。いずれの場合も、指定されたディレクトリの完全パスが計算され、カレントディレクトリとして格納されます。</para>
      /// </param>
      /// <returns>
      ///   <para>関数が成功した場合、戻り値はゼロ以外です。</para>
      ///   <para>関数が失敗した場合、戻り値はゼロです。拡張エラー情報を取得するには GetLastError を呼び出してください。</para>
      /// </returns>
      [SuppressMessage("Microsoft.Usage", "CA2205:UseManagedEquivalentsOfWin32Api")]
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SetCurrentDirectoryW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool SetCurrentDirectory([MarshalAs(UnmanagedType.LPWStr)] string lpPathName);
   }
}
