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
   /// <summary>A KTM transaction object for use with the transacted operations in <see cref="Filesystem"/>.</summary>
   public sealed class KernelTransaction : MarshalByRefObject, IDisposable
   {
      // IKernelTransaction IID: 79427A2B-F895-40e0-BE79-B57DC82ED231
      private static readonly Guid IID_IKernelTransaction = new Guid("79427A2B-F895-40e0-BE79-B57DC82ED231");

      /// <summary>Initializes a new instance of the <see cref="KernelTransaction"/> class, internally using the specified <see cref="Transaction"/>.
      /// This method allows the usage of methods accepting a <see cref="KernelTransaction"/> with an instance of <see cref="System.Transactions.Transaction"/>.
      /// </summary>
      /// <remarks>
      /// <para>This constructor is NOT compatible with NativeAOT deployment because
      /// <see cref="TransactionInterop.GetDtcTransaction"/> and <see cref="Marshal.GetIUnknownForObject"/>
      /// internally rely on built-in COM interop which is not supported in NativeAOT.</para>
      /// <para>For AOT-safe usage, use the other constructors that call CreateTransaction directly.</para>
      /// </remarks>
      /// <param name="transaction">The transaction to use for any transactional operations.</param>
      [SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
      [SecurityCritical]
      [RequiresUnreferencedCode("This constructor uses TransactionInterop.GetDtcTransaction and Marshal.GetIUnknownForObject which require COM interop not available in NativeAOT. Use the parameterless constructor or the constructor with timeout/description parameters instead.")]
      public unsafe KernelTransaction(Transaction transaction)
      {
         var dtcTransaction = TransactionInterop.GetDtcTransaction(transaction);
         var punk = Marshal.GetIUnknownForObject(dtcTransaction);
         try
         {
            // QueryInterface for IKernelTransaction
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
               // Release IKernelTransaction
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)ktxVtable[2];
               releaseFn(pKtx);
            }
         }
         finally
         {
            Marshal.Release(punk);
         }
      }

      /// <summary>Initializes a new instance of the <see cref="KernelTransaction"/> class with a default security descriptor, infinite timeout and no description.</summary>
      [SecurityCritical]
      public KernelTransaction()
         : this(0, null)
      {
      }

      /// <summary>Initializes a new instance of the <see cref="KernelTransaction"/> class with a default security descriptor.</summary>
      /// <param name="timeout"><para>The time, in milliseconds, when the transaction will be aborted if it has not already reached the prepared state.</para></param>
      /// <param name="description">A user-readable description of the transaction. This parameter may be <c>null</c>.</param>
      [SecurityCritical]      
      public KernelTransaction(int timeout, string description)
         : this(null, timeout, description)
      {
      }

      /// <summary>Initializes a new instance of the <see cref="KernelTransaction"/> class.</summary>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
      /// <param name="securityDescriptor">The <see cref="ObjectSecurity"/> security descriptor.</param>
      /// <param name="timeout"><para>The time, in milliseconds, when the transaction will be aborted if it has not already reached the prepared state.</para>
      /// <para>Specify 0 to provide an infinite timeout.</para></param>
      /// <param name="description">A user-readable description of the transaction. This parameter may be <c>null</c>.</param>
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

      /// <summary>Requests that the specified transaction be committed.</summary>
      /// <exception cref="TransactionAlreadyCommittedException"/>
      /// <exception cref="TransactionAlreadyAbortedException"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
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

      /// <summary>Requests that the specified transaction be rolled back. This function is synchronous.</summary>
      /// <exception cref="TransactionAlreadyCommittedException"/>
      /// <exception cref="Win32Exception"/>
      /// <exception cref="PlatformNotSupportedException">The operating system is older than Windows Vista.</exception>
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

      /// <summary>Gets the safe handle.</summary>
      /// <value>The safe handle.</value>
      public SafeHandle SafeHandle
      {
         get { return _hTrans; }
      }

      private readonly SafeKernelTransactionHandle _hTrans;

      #region IDisposable Members

      /// <summary>Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.</summary>
      public void Dispose()
      {
         _hTrans.Close();
      }

      #endregion // IDisposable Members
   }
}
