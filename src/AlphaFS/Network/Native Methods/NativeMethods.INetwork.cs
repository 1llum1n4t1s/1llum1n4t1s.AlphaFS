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
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Alphaleonis.Win32.Network
{
   internal static partial class NativeMethods
   {
      /// <summary>AOT-safe wrapper for the INetwork COM interface (IDispatch-based).
      /// INetwork GUID: DCB00002-570F-4A9B-8D69-199FDBA5723B
      /// Vtable layout: IUnknown (3) + IDispatch (4) + INetwork methods starting at slot 7.
      /// This wrapper takes ownership of the COM pointer passed to the constructor (no AddRef).
      /// The caller must call Dispose() to release the COM reference.</summary>
      internal readonly unsafe struct NetworkWrapper : IDisposable
      {
         private readonly nint _ptr;

         /// <summary>Takes ownership of the COM pointer. Does NOT call AddRef.</summary>
         internal NetworkWrapper(nint comPtr)
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

         /// <summary>Slot 7: GetName - returns BSTR.</summary>
         internal string GetName()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)vtable[7];
            nint bstr;
            int hr = fn(_ptr, &bstr);
            Marshal.ThrowExceptionForHR(hr);
            try { return Marshal.PtrToStringBSTR(bstr); }
            finally { Marshal.FreeBSTR(bstr); }
         }

         /// <summary>Slot 8: SetName.</summary>
         internal void SetName(string name)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint, int>)vtable[8];
            var bstr = Marshal.StringToBSTR(name);
            try { Marshal.ThrowExceptionForHR(fn(_ptr, bstr)); }
            finally { Marshal.FreeBSTR(bstr); }
         }

         /// <summary>Slot 9: GetDescription - returns BSTR.</summary>
         internal string GetDescription()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)vtable[9];
            nint bstr;
            int hr = fn(_ptr, &bstr);
            Marshal.ThrowExceptionForHR(hr);
            try { return Marshal.PtrToStringBSTR(bstr); }
            finally { Marshal.FreeBSTR(bstr); }
         }

         /// <summary>Slot 10: SetDescription.</summary>
         internal void SetDescription(string description)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint, int>)vtable[10];
            var bstr = Marshal.StringToBSTR(description);
            try { Marshal.ThrowExceptionForHR(fn(_ptr, bstr)); }
            finally { Marshal.FreeBSTR(bstr); }
         }

         /// <summary>Slot 11: GetNetworkId - returns GUID.</summary>
         internal Guid GetNetworkId()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, Guid*, int>)vtable[11];
            Guid result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 12: GetDomainType.</summary>
         internal DomainType GetDomainType()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, DomainType*, int>)vtable[12];
            DomainType result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 13: GetNetworkConnections - returns IEnumNetworkConnections (IUnknown ptr).</summary>
         internal IEnumerable<NetworkConnectionWrapper> GetNetworkConnections()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)vtable[13];
            nint enumPtr;
            Marshal.ThrowExceptionForHR(fn(_ptr, &enumPtr));
            try
            {
               return EnumVariantHelper.EnumerateComCollection(enumPtr, ptr => new NetworkConnectionWrapper(ptr));
            }
            finally
            {
               ComHelper.Release(enumPtr);
            }
         }

         /// <summary>Slot 14: GetTimeCreatedAndConnected.</summary>
         internal void GetTimeCreatedAndConnected(out uint pdwLowDateTimeCreated, out uint pdwHighDateTimeCreated, out uint pdwLowDateTimeConnected, out uint pdwHighDateTimeConnected)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, uint*, uint*, uint*, uint*, int>)vtable[14];
            uint a, b, c, d;
            Marshal.ThrowExceptionForHR(fn(_ptr, &a, &b, &c, &d));
            pdwLowDateTimeCreated = a;
            pdwHighDateTimeCreated = b;
            pdwLowDateTimeConnected = c;
            pdwHighDateTimeConnected = d;
         }

         /// <summary>Slot 15: get_IsConnected - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnected
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[15];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 16: get_IsConnectedToInternet - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnectedToInternet
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[16];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 17: GetConnectivity.</summary>
         internal ConnectivityStates GetConnectivity()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, ConnectivityStates*, int>)vtable[17];
            ConnectivityStates result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 18: GetCategory.</summary>
         internal NetworkCategory GetCategory()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, NetworkCategory*, int>)vtable[18];
            NetworkCategory result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }

         /// <summary>Slot 19: SetCategory.</summary>
         internal void SetCategory(NetworkCategory newCategory)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, NetworkCategory, int>)vtable[19];
            Marshal.ThrowExceptionForHR(fn(_ptr, newCategory));
         }
      }
   }
}
