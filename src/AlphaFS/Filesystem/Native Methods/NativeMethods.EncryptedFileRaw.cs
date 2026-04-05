using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   internal partial class NativeMethods
   {
      /// <summary>暗号化されたファイルをバックアップ（エクスポート）または復元（インポート）するために開きます。</summary>
      /// <returns>関数が成功した場合、ERROR_SUCCESS を返します。</returns>
      /// <returns>関数が失敗した場合、WinError.h で定義されたゼロ以外のエラーコードを返します。FORMAT_MESSAGE_FROM_SYSTEM フラグを使用して FormatMessage を呼び出すことで、エラーの一般的なテキスト説明を取得できます。</returns>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      /// <param name="lpFileName">開くファイルの名前。</param>
      /// <param name="ulFlags">実行する操作。</param>
      /// <param name="pvContext">[out] 後続の ReadEncryptedFileRaw、WriteEncryptedFileRaw、または CloseEncryptedFileRaw の呼び出しで
      /// 提示する必要があるコンテキストブロックのアドレス。</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("Advapi32.dll", SetLastError = false, CharSet = CharSet.Unicode, EntryPoint = "OpenEncryptedFileRawW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint OpenEncryptedFileRaw([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, EncryptedFileRawMode ulFlags, out SafeEncryptedFileRawHandle pvContext);


      /// <summary>バックアップまたは復元操作後に暗号化されたファイルを閉じ、関連するシステムリソースを解放します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      /// <param name="pvContext">システム定義のコンテキストブロックへのポインタ。OpenEncryptedFileRaw 関数がコンテキストブロックを返します。</param>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("Advapi32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      internal static extern void CloseEncryptedFileRaw(IntPtr pvContext);


      /// <summary>暗号化されたファイルをバックアップ（エクスポート）します。これは、ファイルを暗号化された状態のまま維持しながらバックアップおよび復元機能を実装するための暗号化ファイルシステム (EFS) 関数グループの1つです。</summary>
      /// <returns>関数が成功した場合、ERROR_SUCCESS を返します。</returns>
      /// <returns>関数が失敗した場合、WinError.h で定義されたゼロ以外のエラーコードを返します。FORMAT_MESSAGE_FROM_SYSTEM フラグを使用して FormatMessage を呼び出すことで、エラーの一般的なテキスト説明を取得できます。</returns>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule"), SuppressUnmanagedCodeSecurity]
      [DllImport("Advapi32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint ReadEncryptedFileRaw([MarshalAs(UnmanagedType.FunctionPtr)] EncryptedFileRawExportCallback pfExportCallback, IntPtr pvCallbackContext, SafeEncryptedFileRawHandle pvContext);


      /// <summary>暗号化されたファイルを復元（インポート）します。これは、ファイルを暗号化された状態のまま維持しながらバックアップおよび復元機能を実装するための暗号化ファイルシステム (EFS) 関数グループの1つです。</summary>
      /// <returns>関数が成功した場合、ERROR_SUCCESS を返します。</returns>
      /// <returns>関数が失敗した場合、WinError.h で定義されたゼロ以外のエラーコードを返します。FORMAT_MESSAGE_FROM_SYSTEM フラグを使用して FormatMessage を呼び出すことで、エラーの一般的なテキスト説明を取得できます。</returns>
      /// <remarks>サポートされる最小クライアント: Windows XP Professional [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("Advapi32.dll", SetLastError = false, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint WriteEncryptedFileRaw([MarshalAs(UnmanagedType.FunctionPtr)] EncryptedFileRawImportCallback pfExportCallback, IntPtr pvCallbackContext, SafeEncryptedFileRawHandle pvContext);


      [SuppressUnmanagedCodeSecurity]
      internal delegate int EncryptedFileRawExportCallback(IntPtr pbData, IntPtr pvCallbackContext, uint ulLength);

      [SuppressUnmanagedCodeSecurity]
      internal delegate int EncryptedFileRawImportCallback(IntPtr pbData, IntPtr pvCallbackContext, ref uint ulLength);
   }
}
