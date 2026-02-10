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
      /// <summary>AOT-safe wrapper for the INetworkListManager COM interface (IDispatch-based).
      /// INetworkListManager GUID: DCB00000-570F-4A9B-8D69-199FDBA5723B
      /// Vtable layout: IUnknown (3) + IDispatch (4) + INetworkListManager methods starting at slot 7.</summary>
      internal sealed unsafe class NetworkListManagerWrapper : IDisposable
      {
         private nint _ptr;

         internal NetworkListManagerWrapper(nint comPtr)
         {
            _ptr = comPtr;
         }

         internal bool IsValid => _ptr != 0;

         public void Dispose()
         {
            var ptr = _ptr;
            _ptr = 0;
            if (ptr != 0)
            {
               nint* vtable = *(nint**)ptr;
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)vtable[2];
               releaseFn(ptr);
            }
         }

         /// <summary>Slot 7: GetNetworks - returns IEnumNetworks (collection).</summary>
         internal IEnumerable<NetworkWrapper> GetNetworks(NetworkConnectivityLevels flags)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, NetworkConnectivityLevels, nint*, int>)vtable[7];
            nint enumPtr;
            Marshal.ThrowExceptionForHR(fn(_ptr, flags, &enumPtr));
            try
            {
               return EnumVariantHelper.EnumerateComCollection(enumPtr, ptr => new NetworkWrapper(ptr));
            }
            finally
            {
               ComHelper.Release(enumPtr);
            }
         }

         /// <summary>Slot 8: GetNetwork - returns INetwork for a given GUID. Caller takes ownership.</summary>
         internal NetworkWrapper GetNetwork(Guid gdNetworkId)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, Guid, nint*, int>)vtable[8];
            nint networkPtr;
            Marshal.ThrowExceptionForHR(fn(_ptr, gdNetworkId, &networkPtr));
            return new NetworkWrapper(networkPtr);
         }

         /// <summary>Slot 9: GetNetworkConnections - returns IEnumNetworkConnections (collection).</summary>
         internal IEnumerable<NetworkConnectionWrapper> GetNetworkConnections()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, nint*, int>)vtable[9];
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

         /// <summary>Slot 10: GetNetworkConnection - returns INetworkConnection for a given GUID. Caller takes ownership.</summary>
         internal NetworkConnectionWrapper GetNetworkConnection(Guid gdNetworkConnectionId)
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, Guid, nint*, int>)vtable[10];
            nint connectionPtr;
            Marshal.ThrowExceptionForHR(fn(_ptr, gdNetworkConnectionId, &connectionPtr));
            return new NetworkConnectionWrapper(connectionPtr);
         }

         /// <summary>Slot 11: get_IsConnected - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnected
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[11];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 12: get_IsConnectedToInternet - property getter returning VARIANT_BOOL.</summary>
         internal bool IsConnectedToInternet
         {
            get
            {
               nint* vtable = *(nint**)_ptr;
               var fn = (delegate* unmanaged[Stdcall]<nint, short*, int>)vtable[12];
               short result;
               Marshal.ThrowExceptionForHR(fn(_ptr, &result));
               return result != 0;
            }
         }

         /// <summary>Slot 13: GetConnectivity.</summary>
         internal ConnectivityStates GetConnectivity()
         {
            nint* vtable = *(nint**)_ptr;
            var fn = (delegate* unmanaged[Stdcall]<nint, ConnectivityStates*, int>)vtable[13];
            ConnectivityStates result;
            Marshal.ThrowExceptionForHR(fn(_ptr, &result));
            return result;
         }
      }
   }
}
