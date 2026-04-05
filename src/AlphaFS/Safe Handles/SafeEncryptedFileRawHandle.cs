using System;
using System.Security;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>OpenEncryptedFileRaw Win32 API関数で使用されるハンドルのラッパークラスを表します。</summary>
   [SecurityCritical]
   internal sealed class SafeEncryptedFileRawHandle : SafeHandleZeroOrMinusOneIsInvalid
   {
      /// <summary>このクラスのデフォルトインスタンスの作成を防止するコンストラクタ。</summary>
      private SafeEncryptedFileRawHandle() : base(true)
      {
      }


      /// <summary><see cref="SafeEncryptedFileRawHandle"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="handle">ハンドル。</param>
      /// <param name="callerHandle">ファイナライズ段階でハンドルを確実に解放する場合は<c>true</c>、確実な解放を防止する場合は<c>false</c>（非推奨）。</param>
      public SafeEncryptedFileRawHandle(IntPtr handle, bool callerHandle) : base(callerHandle)
      {
         SetHandle(handle);
      }

      
      /// <summary>派生クラスでオーバーライドされた場合、ハンドルを解放するために必要なコードを実行します。</summary>
      /// <returns>
      /// ハンドルが正常に解放された場合は<c>true</c>、致命的な障害が発生した場合は
      /// <c>false</c>。この場合、ReleaseHandleFailed マネージデバッグアシスタントが生成されます。
      /// </returns>
      protected override bool ReleaseHandle()
      {
         NativeMethods.CloseEncryptedFileRaw(handle);

         return true;
      }
   }
}
