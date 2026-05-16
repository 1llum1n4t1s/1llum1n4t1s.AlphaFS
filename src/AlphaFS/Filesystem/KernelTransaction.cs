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
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using System.Transactions;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary><see cref="Filesystem"/>のトランザクション操作で使用するKTMトランザクションオブジェクト。</summary>
   public sealed class KernelTransaction : MarshalByRefObject, IDisposable
   {
      // IKernelTransaction IID: 79427A2B-F895-40e0-BE79-B57DC82ED231
      private static readonly Guid IID_IKernelTransaction = new Guid("79427A2B-F895-40e0-BE79-B57DC82ED231");

      /// <summary>指定された<see cref="Transaction"/>を内部的に使用して、<see cref="KernelTransaction"/>クラスの新しいインスタンスを初期化します。
      /// このメソッドにより、<see cref="System.Transactions.Transaction"/>のインスタンスで<see cref="KernelTransaction"/>を受け入れるメソッドを使用できます。
      /// </summary>
      /// <remarks>
      /// <para>このコンストラクタはNativeAOTデプロイメントと互換性がありません。
      /// <see cref="TransactionInterop.GetDtcTransaction"/>と<see cref="Marshal.GetIUnknownForObject"/>は
      /// 内部的にNativeAOTでサポートされていない組み込みCOM相互運用に依存しているためです。</para>
      /// <para>AOT安全な使用方法については、CreateTransactionを直接呼び出す他のコンストラクタを使用してください。</para>
      /// </remarks>
      /// <param name="transaction">トランザクション操作に使用するトランザクション。</param>
      [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
      [SecurityCritical]
      [RequiresUnreferencedCode("This constructor uses TransactionInterop.GetDtcTransaction and Marshal.GetIUnknownForObject which require COM interop not available in NativeAOT. Use the parameterless constructor or the constructor with timeout/description parameters instead.")]
      public unsafe KernelTransaction(Transaction transaction)
      {
         var dtcTransaction = TransactionInterop.GetDtcTransaction(transaction);
         var punk = Marshal.GetIUnknownForObject(dtcTransaction);
         try
         {
            // IKernelTransactionのQueryInterface
            var iid = IID_IKernelTransaction;
            nint* vtable = *(nint**)punk;
            var queryFn = (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)vtable[0];
            nint pKtx;
            Marshal.ThrowExceptionForHR(queryFn(punk, &iid, &pKtx));

            // IKernelTransaction vtable: IUnknown(3) + [3] GetHandle
            nint* ktxVtable = *(nint**)pKtx;
            try
            {
               var getHandleFn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)ktxVtable[3];
               nint rawHandle;
               Marshal.ThrowExceptionForHR(getHandleFn(pKtx, &rawHandle));
               _hTrans = new SafeKernelTransactionHandle(rawHandle);
            }
            finally
            {
               // IKernelTransactionを解放
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)ktxVtable[2];
               releaseFn(pKtx);
            }
         }
         finally
         {
            Marshal.Release(punk);
         }
      }

      /// <summary>デフォルトのセキュリティ記述子、無限のタイムアウト、説明なしで<see cref="KernelTransaction"/>クラスの新しいインスタンスを初期化します。</summary>
      [SecurityCritical]
      public KernelTransaction()
         : this(0, null)
      {
      }

      /// <summary>デフォルトのセキュリティ記述子で<see cref="KernelTransaction"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="timeout"><para>トランザクションが準備状態に達していない場合に中止されるまでの時間（ミリ秒）。</para></param>
      /// <param name="description">ユーザーが読み取り可能なトランザクションの説明。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]      
      public KernelTransaction(int timeout, string description)
         : this(null, timeout, description)
      {
      }

      /// <summary><see cref="KernelTransaction"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <exception cref="PlatformNotSupportedException">オペレーティングシステムがWindows Vistaより古い場合。</exception>
      /// <param name="securityDescriptor"><see cref="ObjectSecurity"/>セキュリティ記述子。</param>
      /// <param name="timeout"><para>トランザクションが準備状態に達していない場合に中止されるまでの時間（ミリ秒）。</para>
      /// <para>無限のタイムアウトを指定するには0を指定します。</para></param>
      /// <param name="description">ユーザーが読み取り可能なトランザクションの説明。このパラメータは<c>null</c>にできます。</param>
      [SecurityCritical]
      public KernelTransaction(ObjectSecurity securityDescriptor, int timeout, string description)
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }

         using var securityAttributes = new Security.NativeMethods.SecurityAttributes(securityDescriptor);
         _hTrans = NativeMethods.CreateTransaction(securityAttributes, IntPtr.Zero, 0, 0, 0, timeout, description);
         var lastError = Marshal.GetLastWin32Error();            

         NativeMethods.IsValidHandle(_hTrans, lastError);
      }

      /// <summary>指定されたトランザクションのコミットを要求します。</summary>
      /// <exception cref="TransactionAlreadyCommittedException"/>
      /// <exception cref="TransactionAlreadyAbortedException"/>
      /// <exception cref="PlatformNotSupportedException">オペレーティングシステムがWindows Vistaより古い場合。</exception>
      /// <exception cref="Win32Exception"/>
      [SecurityCritical]
      public void Commit()
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }

         if (!NativeMethods.CommitTransaction(_hTrans))
         {
            CheckTransaction();
         }
      }

      /// <summary>指定されたトランザクションのロールバックを要求します。この関数は同期的です。</summary>
      /// <exception cref="TransactionAlreadyCommittedException"/>
      /// <exception cref="Win32Exception"/>
      /// <exception cref="PlatformNotSupportedException">オペレーティングシステムがWindows Vistaより古い場合。</exception>
      [SecurityCritical]
      public void Rollback()
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int) Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }

         if (!NativeMethods.RollbackTransaction(_hTrans))
         {
            CheckTransaction();
         }
      }

      private static void CheckTransaction()
      {
         var error = (uint) Marshal.GetLastWin32Error();
         var hr = Marshal.GetHRForLastWin32Error();

         switch (error)
         {
            case Win32Errors.ERROR_TRANSACTION_ALREADY_ABORTED:
               throw new TransactionAlreadyAbortedException("Transaction was already aborted", Marshal.GetExceptionForHR(hr));

            case Win32Errors.ERROR_TRANSACTION_ALREADY_COMMITTED:
               throw new TransactionAlreadyCommittedException("Transaction was already committed", Marshal.GetExceptionForHR(hr));

            default:
               Marshal.ThrowExceptionForHR(Marshal.GetHRForLastWin32Error());
               break;
         }
      }

      /// <summary>セーフハンドルを取得します。</summary>
      /// <value>セーフハンドル。</value>
      public SafeHandle SafeHandle
      {
         get { return _hTrans; }
      }

      private readonly SafeKernelTransactionHandle _hTrans;

      #region IDisposable Members

      /// <summary>アンマネージリソースの解放、リリース、またはリセットに関連するアプリケーション定義のタスクを実行します。</summary>
      public void Dispose()
      {
         _hTrans.Close();
      }

      #endregion // IDisposable Members
   }
}
