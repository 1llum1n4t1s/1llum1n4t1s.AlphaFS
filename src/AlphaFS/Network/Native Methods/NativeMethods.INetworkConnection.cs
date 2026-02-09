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
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Network
{
   internal static partial class NativeMethods
   {
      /// <summary>AOT-safe wrapper for the INetworkConnection COM interface (IDispatch-based).
      /// INetworkConnection GUID: DCB00005-570F-4A9B-8D69-199FDBA5723B
      /// Vtable layout: IUnknown (3) + IDispatch (4) + INetworkConnection methods starting at slot 7.
      /// This wrapper takes ownership of the COM pointer passed to the constructor (no AddRef).
      /// The caller must call Dispose() to release the COM reference.</summary>
      internal readonly unsafe struct NetworkConnectionWrapper : IDisposable
      {
         private readonly nint _ptr;

         /// <summary>Takes ownership of the COM pointer. Does NOT call AddRef.</summary>
         internal NetworkConnectionWrapper(nint comPtr)
         {
            _ptr = comPtr;
         }

         internal bool IsValid => _ptr != 0;

         /// <summary>Releases the COM reference.</summary>
         public void Dispose()
         {
            if (_ptr != 0)
            {
               nint* vtable = *(nint**)_ptr;
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)vtable[2];
               releaseFn(_ptr);
            }
         }

         /// <summary>Slot 7: GetNetwork - returns INetwork ptr. Caller takes ownership.</summary>
         internal NetworkWrapper GetNetwork()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)vtable[7];
            nint networkPtr;
            Marshal.ThrowExceptionForHR(fn(_ptr, &networkPtr));
            return new NetworkWrapper(networkPtr);
         }

         /// <summary>Slot 8: get_IsConnected - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnected
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[8];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 9: get_IsConnectedToInternet - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnectedToInternet
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[9];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 10: GetConnectivity.</summary>
         internal ConnectivityStates GetConnectivity()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, ConnectivityStates*, int>)vtable[10];
            ConnectivityStates result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 11: GetConnectionId.</summary>
         internal Guid GetConnectionId()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, Guid*, int>)vtable[11];
            Guid result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 12: GetAdapterId.</summary>
         internal Guid GetAdapterId()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, Guid*, int>)vtable[12];
            Guid result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 13: GetDomainType.</summary>
         internal DomainType GetDomainType()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, DomainType*, int>)vtable[13];
            DomainType result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }
      }
   }
}
