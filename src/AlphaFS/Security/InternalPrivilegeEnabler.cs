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
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;

namespace Alphaleonis.Win32.Security
{
   /// <summary>
   /// このオブジェクトは、現在実行中のプロセスの存続期間中に特定の特権を有効にするために使用されます。
   /// 昇格された特権が不要になったらすぐにDisposeする必要があります。
   /// 詳細については、MSDNのAdjustTokenPrivilegesのドキュメントを参照してください。
   /// </summary>
   internal sealed class InternalPrivilegeEnabler : IDisposable
   {
      /// <summary><see cref="PrivilegeEnabler"/>クラスの新しいインスタンスを初期化し、現在実行中のプロセスに対して指定された特権を有効にします。</summary>
      /// <param name="privilegeName">特権の名前。</param>
      [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
      [SecurityCritical]
      public InternalPrivilegeEnabler(Privilege privilegeName)
      {
         if (null == privilegeName)
         {
            throw new ArgumentNullException("privilegeName");
         }

         EnabledPrivilege = privilegeName;
         AdjustPrivilege(true);
      }


      /// <summary>
      /// アンマネージリソースの解放、リリース、またはリセットに関連するアプリケーション定義のタスクを実行します。
      /// この場合、以前に有効にされた特権が無効になります。
      /// </summary>            
      public void Dispose()
      {
         try
         {
            if (null != EnabledPrivilege)
            {
               AdjustPrivilege(false);
            }
         }
         finally
         {
            EnabledPrivilege = null;
         }
      }


      /// <summary>SE_PRIVILEGE_ENABLED: TOKEN_PRIVILEGES.Attributes の「特権が有効」フラグ。</summary>
      private const uint SePrivilegeEnabled = 2;


      public Privilege EnabledPrivilege { get; private set; }


      /// <summary>特権を調整します。</summary>
      /// <param name="enable"><c>true</c>の場合、特権が有効になります。それ以外の場合は無効になります。</param>
      [SecurityCritical]
      private void AdjustPrivilege(bool enable)
      {
         using var currentIdentity = WindowsIdentity.GetCurrent(TokenAccessLevels.Query | TokenAccessLevels.AdjustPrivileges);
         uint length;
         var hToken = currentIdentity.Token;
         var mOldPrivilege = new TOKEN_PRIVILEGES();

         var newPrivilege = new TOKEN_PRIVILEGES
         {
            PrivilegeCount = 1,
            Luid = Filesystem.NativeMethods.LongToLuid(EnabledPrivilege.LookupLuid()),

            Attributes = enable ? SePrivilegeEnabled : 0
         };


         var success = NativeMethods.AdjustTokenPrivileges(hToken, false, ref newPrivilege, (uint) Marshal.SizeOf(mOldPrivilege), out mOldPrivilege, out length);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError);
         }

         // MSDN AdjustTokenPrivileges 仕様: success=TRUE であっても GetLastError が ERROR_NOT_ALL_ASSIGNED (1300) を返す場合、
         // 要求した特権の一部または全部がトークンに割り当てられておらず実際には有効化されていない。
         // このサイレント失敗を検出して呼び出し側に通知することで、後続のファイル操作の謎の権限エラーを防止する。
         // 無効化 (enable == false) の場合、特権が既にトークンに無いだけのため警告対象外。
         if (enable && lastError == Win32Errors.ERROR_NOT_ALL_ASSIGNED)
         {
            // 特権名はフィールドをクリアする前に取り出す。以前は EnabledPrivilege = null を先に実行していたため、
            // 続くメッセージ組み立てで EnabledPrivilege?.ToString() が常に null に評価され、
            // どの特権が不足しているのか分からない LUID 構造体の型名だけが表示されていた。
            var privilegeName = EnabledPrivilege?.ToString() ?? newPrivilege.Luid.ToString();

            EnabledPrivilege = null;

            throw new UnauthorizedAccessException(
               string.Format(System.Globalization.CultureInfo.CurrentCulture,
                  "特権 '{0}' をトークンに割り当てられませんでした。プロセスがこの特権を保持していない可能性があります（管理者権限が必要な場合があります）。",
                  privilegeName));
         }

         // 巻き戻しが必要かどうかを判定する。EnabledPrivilege を null にすると Dispose 時の
         // AdjustPrivilege(false) が実行されなくなる。
         //
         // PrivilegeCount == 0 だけでは不十分。AdjustTokenPrivileges は NewState に列挙した特権を
         // 「値が変わったかどうかに関わらず」PreviousState へ返すため、元から有効だった特権でも
         // PrivilegeCount は 1 になる。前の状態が SE_PRIVILEGE_ENABLED だったかは Attributes でしか
         // 判別できず、これを見ないとスコープ離脱時にホストプロセスが常用している特権まで無効化してしまう
         // (トークンはプロセス共有なので他スレッドにも波及する)。
         // PrivilegeEnabler の公開ドキュメント「既に有効な特権は無効にされません。」はこの挙動を約束している。
         if (mOldPrivilege.PrivilegeCount == 0 || (mOldPrivilege.Attributes & SePrivilegeEnabled) != 0)
         {
            EnabledPrivilege = null;
         }
      }
   }
}
