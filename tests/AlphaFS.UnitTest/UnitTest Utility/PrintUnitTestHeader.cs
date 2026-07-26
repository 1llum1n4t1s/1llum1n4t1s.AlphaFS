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
