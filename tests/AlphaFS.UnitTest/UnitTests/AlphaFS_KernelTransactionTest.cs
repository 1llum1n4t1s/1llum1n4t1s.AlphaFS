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

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AlphaFS.UnitTest
{
   /// <summary>トランザクショナル NTFS (TxF) を使う *Transacted API のスモークテスト。</summary>
   /// <remarks>
   ///   AlphaFS の公開 API は「通常版」と「KernelTransaction を取る Transacted 版」が対になっており、
   ///   Transacted 専用ファイルだけで 100 を超える。ここではその経路が実際に動作すること
   ///   (トランザクション内の変更が外から見えず、commit で反映され、rollback で消えること) を確認する。
   ///
   ///   TxF は Microsoft により非推奨とされているが、公開 API として提供している以上、
   ///   通常版の変更が Transacted 版を壊していないことは検知できる必要がある。
   /// </remarks>
   [TestClass]
   public partial class AlphaFS_KernelTransactionTest
   {
   }
}
