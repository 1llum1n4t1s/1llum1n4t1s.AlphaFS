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
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace Alphaleonis.Win32.Security
{
   /// <summary>[AlphaFS] 現在のプロセスのコンテキストを判定するクラス。</summary>
   public static class ProcessContext
   {
      #region Properties

      /// <summary>[AlphaFS] 現在のプロセスが管理者のコンテキストで実行されているかどうかを判定します。</summary>
      /// <returns>現在のプロセスが管理者のコンテキストで実行されている場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public static bool IsAdministrator
      {
         get
         {
            var principal = GetWindowsPrincipal(out var windowsIdentity);

            using (windowsIdentity)
               return

                  // ローカル管理者
                  principal.IsInRole(WindowsBuiltInRole.Administrator) ||

                  // ドメイン管理者
                  principal.IsInRole(512);
         }
      }


      /// <summary>[AlphaFS] UACが有効で、現在のプロセスが昇格された状態にあるかどうかを判定します。
      /// <para>現在のユーザーがデフォルトのAdministratorの場合、プロセスは昇格された状態にあると見なされます。</para>
      /// <para>これは、デフォルトのAdministrator（デフォルトでは無効）がUACプロンプトを表示せずにすべてのアクセス権を取得するためです。</para>
      /// </summary>
      /// <returns>UACが有効で現在のプロセスが昇格された状態にある場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public static bool IsElevatedProcess
      {
         get
         {
            return IsUacEnabled && (GetProcessElevationType() == NativeMethods.TOKEN_ELEVATION_TYPE.TokenElevationTypeFull || IsAdministrator);
         }
      }


      /// <summary>[AlphaFS] ローカルコンピューターの"EnableLUA"レジストリキーを読み取ることでUACが有効かどうかを判定します。</summary>
      /// <returns>UACのステータスがレジストリから正常に読み取られた場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Uac")]
      public static bool IsUacEnabled
      {
         get
         {
            using var uacKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Policies\System", false);
            var enableLua = uacKey?.GetValue("EnableLUA");
            return null != enableLua && enableLua.Equals(1);
         }
      }


      /// <summary>[AlphaFS] 現在のプロセスがWindowsサービスのコンテキストで実行されているかどうかを判定します。</summary>
      /// <returns>現在のプロセスがWindowsサービスのコンテキストで実行されている場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public static bool IsWindowsService
      {
         get
         {
            var principal = GetWindowsPrincipal(out var windowsIdentity);

            using (windowsIdentity)
               return principal.IsInRole(new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null)) ||
                      principal.IsInRole(new SecurityIdentifier(WellKnownSidType.ServiceSid, null));
         }
      }

      #endregion // Properties


      private static WindowsPrincipal GetWindowsPrincipal(out WindowsIdentity windowsIdentity)
      {
         windowsIdentity = WindowsIdentity.GetCurrent();

         if (null == windowsIdentity)
         {
            throw new InvalidOperationException(Resources.GetCurrentWindowsIdentityFailed);
         }

         return new WindowsPrincipal(windowsIdentity);
      }


      /// <summary>[AlphaFS] 現在のプロセスの昇格タイプを取得します。</summary>
      /// <returns><see cref="NativeMethods.TOKEN_ELEVATION_TYPE"/>の値。</returns>
      [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "GetTokenInformation")]
      [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly", MessageId = "OpenProcessToken")]
      private static NativeMethods.TOKEN_ELEVATION_TYPE GetProcessElevationType()
      {

         bool success;
         SafeTokenHandle tokenHandle;

         // Process インスタンスを保持せずに .Handle だけ渡すと、ネイティブ呼び出しの前に Process が
         // 回収されてハンドルが閉じられ得る。using で呼び出し完了まで生存させる。
         using (var process = Process.GetCurrentProcess())
         {
            success = NativeMethods.OpenProcessToken(process.Handle, NativeMethods.TOKEN.TOKEN_READ, out tokenHandle);
         }

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            throw new Win32Exception(lastError, string.Format(CultureInfo.CurrentCulture, "{0}: OpenProcessToken failed with error: {1}", nameof(GetProcessElevationType), lastError.ToString(CultureInfo.CurrentCulture)));
         }


         using (tokenHandle)
         using (var safeBuffer = new SafeGlobalMemoryBufferHandle(Marshal.SizeOf<int>()))
         {
            uint bytesReturned;
            success = NativeMethods.GetTokenInformation(tokenHandle, NativeMethods.TOKEN_INFORMATION_CLASS.TokenElevationType, safeBuffer, (uint) safeBuffer.Capacity, out bytesReturned);

            lastError = Marshal.GetLastWin32Error();

            if (!success)
            {
               throw new Win32Exception(lastError, string.Format(CultureInfo.CurrentCulture, "{0}: GetTokenInformation failed with error: {1}", nameof(GetProcessElevationType), lastError.ToString(CultureInfo.CurrentCulture)));
            }


            return (NativeMethods.TOKEN_ELEVATION_TYPE) safeBuffer.ReadInt32();
         }
      }
   }
}
