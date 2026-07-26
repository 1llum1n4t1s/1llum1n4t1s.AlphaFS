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
using System.Globalization;
using System.Security.AccessControl;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   /// <summary>Containts static variables, used by unit tests.</summary>
   public static partial class UnitTestConstants
   {
      /// <summary>ネットワーク側の検証を明示的に無効化する環境変数。CI のように SMB 管理共有へ到達できない環境で 1 / true を設定する。</summary>
      public const string SkipNetworkTestsVariable = "ALPHAFS_SKIP_NETWORK_TESTS";

      private static readonly Lazy<bool> NetworkShareAvailable = new Lazy<bool>(DetectNetworkShare);


      /// <summary>UNC (管理共有) 経由の検証が実行できる環境かどうか。</summary>
      public static bool IsNetworkTestingAvailable
      {
         get { return NetworkShareAvailable.Value; }
      }


      private static bool DetectNetworkShare()
      {
         var skip = Environment.GetEnvironmentVariable(SkipNetworkTestsVariable);

         if (!string.IsNullOrWhiteSpace(skip) && !skip.Equals("0", StringComparison.OrdinalIgnoreCase) && !skip.Equals("false", StringComparison.OrdinalIgnoreCase))
         {
            return false;
         }

         // 実際に管理共有へ到達できるかを 1 度だけ確認する。到達できない環境ではネットワーク側の検証を skip する。
         try
         {
            return System.IO.Directory.Exists(Alphaleonis.Win32.Filesystem.Path.LocalToUnc(Environment.SystemDirectory));
         }
         catch (Exception)
         {
            return false;
         }
      }


      /// <summary>
      ///   管理者権限が必要な検証を、非特権環境では失敗ではなく skip として扱う。
      /// </summary>
      /// <remarks>
      ///   ACL の拒否設定、ボリュームシャドウコピー、タイムスタンプ変更などは昇格した状態でしか成功しない。
      ///   CI runner や通常のユーザーセッションでは実行できないため、リグレッションと区別できるよう skip する。
      /// </remarks>
      public static void RequireElevation(string what)
      {
         if (!Alphaleonis.Win32.Security.ProcessContext.IsElevatedProcess)
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture, "管理者権限が必要なため {0} の検証を skip しました。", what));
         }
      }


      /// <summary>マシン全体の状態を恒久的に変更する検証を有効化する環境変数。既定では実行しない。</summary>
      public const string EnableMachineStateTestsVariable = "ALPHAFS_ENABLE_MACHINE_STATE_TESTS";


      /// <summary>
      ///   マシン全体の状態を恒久的に変更し、後始末できない検証を skip する。
      /// </summary>
      /// <remarks>
      ///   例: MoveOptions.DelayUntilReboot は HKLM の PendingFileRenameOperations にエントリを追加し、
      ///   次回起動時に実際の削除が走る。テストからは取り消せないため、明示的に
      ///   <see cref="EnableMachineStateTestsVariable"/> を設定した場合だけ実行する。
      /// </remarks>
      public static void RequireMachineStateOptIn(string what)
      {
         var enabled = Environment.GetEnvironmentVariable(EnableMachineStateTestsVariable);

         if (string.IsNullOrWhiteSpace(enabled) || enabled.Equals("0", StringComparison.OrdinalIgnoreCase) || enabled.Equals("false", StringComparison.OrdinalIgnoreCase))
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture,
               "{0} はマシン全体の状態を恒久的に変更し取り消せないため skip しました。実行するには環境変数 {1} を 1 に設定してください。", what, EnableMachineStateTestsVariable));
         }
      }


      /// <summary>
      ///   トランザクショナル NTFS (TxF) が使えない環境では skip する。
      /// </summary>
      /// <remarks>
      ///   TxF は NTFS でのみ動作する。ReFS / FAT32 / exFAT、およびネットワークドライブでは使えないため、
      ///   リグレッションと区別できるよう skip する。
      ///   判定には被テスト対象ではなく <see cref="System.IO.DriveInfo"/> を使い、
      ///   AlphaFS 側の回帰が「環境非対応」に化けないようにしている。
      /// </remarks>
      /// <param name="path">検証で使うローカルパス。</param>
      public static void RequireTransactionalNtfs(string path)
      {
         var root = System.IO.Path.GetPathRoot(System.IO.Path.GetFullPath(path));

         string format;

         try
         {
            var drive = new System.IO.DriveInfo(root);

            format = drive.IsReady ? drive.DriveFormat : null;
         }
         catch (Exception ex)
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture,
               "ボリューム [{0}] のファイルシステムを判定できなかったため TxF の検証を skip しました: {1}", root, ex.Message));

            return;
         }

         if (!string.Equals(format, "NTFS", StringComparison.OrdinalIgnoreCase))
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture,
               "ボリューム [{0}] は {1} で、トランザクショナル NTFS (TxF) が使えないため skip しました。", root, format ?? "不明"));
         }
      }


      private static readonly Lazy<DenyAclProbeResult> DenyAclRoundTrip = new Lazy<DenyAclProbeResult>(DetectDenyAclSupport);


      /// <summary>拒否 ACE の往復が可能かどうかと、不可能だった場合の理由。</summary>
      private sealed class DenyAclProbeResult
      {
         public DenyAclProbeResult(bool supported, string reason)
         {
            Supported = supported;
            Reason = reason;
         }

         public bool Supported { get; }

         /// <summary>skip する場合の理由。握り潰した例外の内容を含む。</summary>
         public string Reason { get; }
      }


      /// <summary>
      ///   現在のユーザーに対する拒否 ACE を「付与して解除できる」環境かどうかを実地に確認する。
      /// </summary>
      /// <remarks>
      ///   拒否 ACE は所有者以外に対して WRITE_DAC も拒否するため、作成したオブジェクトの所有者が
      ///   自分自身にならない環境 (昇格プロセスでは所有者が BUILTIN\Administrators になる等) では
      ///   一度付与すると解除できず、テストが後始末できないまま失敗する。
      ///   実際に往復できるかを 1 度だけ試し、できない環境では該当テストを skip する。
      ///
      ///   <para>判定には <see cref="System.IO"/> の ACL API を使い、被テスト対象である
      ///   <c>Alphaleonis.Win32.Filesystem.Directory.GetAccessControl</c> /
      ///   <c>SetAccessControl</c> は使わない。被テスト対象でプローブすると、AlphaFS 側の回帰が
      ///   「環境が非対応」に化けて該当テストが黙って skip され、回帰を見逃す。</para>
      ///
      ///   <para>skip する場合は握り潰した例外の内容をメッセージに含める。理由が残らないと、
      ///   CI 上で skip されたときに環境要因なのか実装の問題なのかを後から判断できない。</para>
      /// </remarks>
      public static void RequireDenyAclRoundTrip(string what)
      {
         var probe = DenyAclRoundTrip.Value;

         if (!probe.Supported)
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture, "拒否 ACE の付与と解除ができない環境のため {0} の検証を skip しました。理由: {1}", what, probe.Reason));
         }
      }


      private static DenyAclProbeResult DetectDenyAclSupport()
      {
         var probe = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "AlphaFS.DenyAclProbe." + Guid.NewGuid().ToString("N").Substring(0, 8));

         System.IO.DirectoryInfo probeInfo;

         try
         {
            probeInfo = System.IO.Directory.CreateDirectory(probe);
         }
         catch (Exception ex)
         {
            return new DenyAclProbeResult(false, string.Format(CultureInfo.CurrentCulture, "プローブ用ディレクトリを作成できませんでした ({0}: {1})", ex.GetType().Name, ex.Message));
         }

         var user = (Environment.UserDomainName + @"\" + Environment.UserName).TrimStart('\\');
         var rule = new FileSystemAccessRule(user, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Deny);

         try
         {
            var security = System.IO.FileSystemAclExtensions.GetAccessControl(probeInfo);
            security.AddAccessRule(rule);
            System.IO.FileSystemAclExtensions.SetAccessControl(probeInfo, security);

            // 付与できても解除できなければ意味が無いので、往復できることまで確認する。
            security = System.IO.FileSystemAclExtensions.GetAccessControl(probeInfo);
            security.RemoveAccessRule(rule);
            System.IO.FileSystemAclExtensions.SetAccessControl(probeInfo, security);

            return new DenyAclProbeResult(true, null);
         }
         catch (Exception ex)
         {
            return new DenyAclProbeResult(false, string.Format(CultureInfo.CurrentCulture, "System.IO でも拒否 ACE を往復できませんでした ({0}: {1})", ex.GetType().Name, ex.Message));
         }
         finally
         {
            try
            {
               System.IO.Directory.Delete(probe, true);
            }
            catch (Exception)
            {
               // 拒否 ACE が外せなかった場合は削除もできない。一時ディレクトリなので放置してよい。
            }
         }
      }


      /// <summary>
      ///   ネットワーク共有 (SMB 管理共有 / ドライブ割り当て) が前提の検証を、到達できない環境では skip として扱う。
      /// </summary>
      public static void RequireNetworkTesting(string what)
      {
         if (!IsNetworkTestingAvailable)
         {
            Assert.Inconclusive(string.Format(CultureInfo.CurrentCulture, "ネットワーク共有へ到達できないため {0} の検証を skip しました。", what));
         }
      }


      public static void PrintUnitTestHeader(bool? isNetwork = null)
      {
         if (null == isNetwork)
         {
            Console.WriteLine("\n=== TEST LOCAL / NETWORK ===");
         }

         else
         {
            Console.WriteLine("\n=== TEST {0} ===", (bool) isNetwork ? "NETWORK" : "LOCAL");
         }


         Console.WriteLine();


         // ネットワーク共有へ到達できない環境では、失敗ではなく skip として扱う。
         // ローカル側の検証は先に実行済みなので、本物のリグレッションは引き続き検知できる。
         if (isNetwork.HasValue && isNetwork.Value && !IsNetworkTestingAvailable)
         {
            Assert.Inconclusive("ネットワーク共有 (SMB 管理共有) へ到達できないため、ネットワーク側の検証を skip しました。");
         }
      }
   }
}
