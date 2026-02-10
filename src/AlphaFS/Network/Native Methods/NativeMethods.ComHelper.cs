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
   /// <summary>AOT-safe COM helper utilities.</summary>
   internal static unsafe class ComHelper
   {
      private static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

      [DllImport("ole32.dll")]
      private static extern int CoCreateInstance(
         ref Guid rclsid,
         nint pUnkOuter,
         uint dwClsContext,
         ref Guid riid,
         out nint ppv);

      private const uint CLSCTX_INPROC_SERVER = 1;
      private const uint CLSCTX_ALL = CLSCTX_INPROC_SERVER | 0x4 | 0x10;

      /// <summary>Creates a COM object instance using CoCreateInstance (AOT-safe).</summary>
      internal static nint CoCreateInstance(Guid clsid, Guid iid)
      {
         int hr = CoCreateInstance(ref clsid, 0, CLSCTX_ALL, ref iid, out var ppv);
         Marshal.ThrowExceptionForHR(hr);
         return ppv;
      }

      /// <summary>Calls IUnknown::Release on a COM pointer.</summary>
      internal static void Release(nint comPtr)
      {
         if (comPtr != 0)
         {
            nint* vtable = *(nint**)comPtr;
            var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)vtable[2];
            releaseFn(comPtr);
         }
      }

      /// <summary>Calls IUnknown::QueryInterface on a COM pointer.</summary>
      internal static nint QueryInterface(nint comPtr, Guid iid)
      {
         nint* vtable = *(nint**)comPtr;
         var queryFn = (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)vtable[0];
         nint result;
         int hr = queryFn(comPtr, &iid, &result);
         Marshal.ThrowExceptionForHR(hr);
         return result;
      }
   }

   /// <summary>AOT-safe IEnumVARIANT helper for enumerating COM collections.</summary>
   internal static unsafe class EnumVariantHelper
   {
      // IEnumVARIANT IID
      private static readonly Guid IID_IEnumVARIANT = new Guid("00020404-0000-0000-C000-000000000046");

      // VARIANT structure size (on 64-bit: 24 bytes)
      private const int VARIANT_SIZE = 24;

      // VT_UNKNOWN = 13, VT_DISPATCH = 9
      private const ushort VT_UNKNOWN = 13;
      private const ushort VT_DISPATCH = 9;

      /// <summary>Enumerates a COM collection object via _NewEnum/IEnumVARIANT.
      /// The collection is expected to support IDispatch with DISPID_NEWENUM (-4) returning an IEnumVARIANT,
      /// OR to directly support IEnumVARIANT.
      /// The wrapItem function receives ownership of each COM pointer (no AddRef/Release needed).
      /// On error, all already-collected items are disposed to prevent COM reference leaks.
      /// Returns all items as a list.</summary>
      internal static List<T> EnumerateComCollection<T>(nint collectionPtr, Func<nint, T> wrapItem) where T : IDisposable
      {
         var result = new List<T>();

         nint enumVariantPtr = GetEnumVariant(collectionPtr);
         if (enumVariantPtr == 0)
            return result;

         try
         {
            // IEnumVARIANT vtable: IUnknown (3) + [3] Next, [4] Skip, [5] Reset, [6] Clone
            nint* vtable = *(nint**)enumVariantPtr;
            var nextFn = (delegate* unmanaged[Stdcall]<nint, uint, nint, uint*, int>)vtable[3];

            // Allocate a VARIANT on the stack
            byte* variantBuf = stackalloc byte[VARIANT_SIZE];

            while (true)
            {
               // Zero the variant
               new Span<byte>(variantBuf, VARIANT_SIZE).Clear();

               uint fetched;
               int hr = nextFn(enumVariantPtr, 1, (nint)variantBuf, &fetched);

               // S_FALSE (1) with fetched==0 means end of enumeration
               if (fetched == 0)
                  break;

               // Real error: throw after cleaning up the VARIANT and already-collected items
               if (hr < 0)
               {
                  ushort vtErr = *(ushort*)variantBuf;
                  nint punkErr = *(nint*)(variantBuf + 8);
                  if ((vtErr == VT_UNKNOWN || vtErr == VT_DISPATCH) && punkErr != 0)
                     ComHelper.Release(punkErr);

                  // Dispose all already-collected items to prevent leaks
                  foreach (var item in result)
                     item.Dispose();

                  Marshal.ThrowExceptionForHR(hr);
               }

               // Read the VARIANT type (first 2 bytes)
               ushort vt = *(ushort*)variantBuf;

               // The COM object pointer is at offset 8 in the VARIANT
               nint punkItem = *(nint*)(variantBuf + 8);

               if ((vt == VT_UNKNOWN || vt == VT_DISPATCH) && punkItem != 0)
               {
                  // Transfer ownership: wrapItem takes the VARIANT's reference directly.
                  // If wrapItem or Add fails, release the COM pointer to prevent leaks.
                  try
                  {
                     result.Add(wrapItem(punkItem));
                  }
                  catch
                  {
                     ComHelper.Release(punkItem);
                     throw;
                  }
               }
            }
         }
         catch
         {
            // On any exception (including from wrapItem), dispose all already-collected items
            foreach (var item in result)
               item.Dispose();

            throw;
         }
         finally
         {
            ComHelper.Release(enumVariantPtr);
         }

         return result;
      }

      /// <summary>Gets an IEnumVARIANT from a collection COM object.
      /// First tries QueryInterface for IEnumVARIANT directly,
      /// then tries IDispatch::Invoke with DISPID_NEWENUM.</summary>
      private static nint GetEnumVariant(nint collectionPtr)
      {
         // First try: QueryInterface for IEnumVARIANT directly
         nint* vtable = *(nint**)collectionPtr;
         var queryFn = (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)vtable[0];
         var iid = IID_IEnumVARIANT;
         nint enumPtr;

         if (queryFn(collectionPtr, &iid, &enumPtr) == 0)
            return enumPtr;

         // Second try: Use IDispatch::Invoke with DISPID_NEWENUM (-4) to get the enumerator
         return GetEnumVariantViaDispatch(collectionPtr);
      }

      /// <summary>Gets IEnumVARIANT via IDispatch::Invoke(DISPID_NEWENUM).</summary>
      private static nint GetEnumVariantViaDispatch(nint dispatchPtr)
      {
         // IDispatch vtable slot 6 = Invoke
         // DISPID_NEWENUM = -4
         // We need to call Invoke with DISPATCH_PROPERTYGET to get the _NewEnum property

         nint* vtable = *(nint**)dispatchPtr;

         // IDispatch::Invoke signature:
         // HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
         var invokeFn = (delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, nint, nint, nint, nint, int>)vtable[6];

         const int DISPID_NEWENUM = -4;
         const ushort DISPATCH_PROPERTYGET = 2;

         var iidNull = Guid.Empty;

         // Allocate DISPPARAMS on stack (32 bytes: 2 pointers + 2 UINTs)
         // typedef struct tagDISPPARAMS { VARIANTARG *rgvarg; DISPID *rgdispidNamedArgs; UINT cArgs; UINT cNamedArgs; } DISPPARAMS;
         byte* dispParamsBuf = stackalloc byte[32];
         *(nint*)(dispParamsBuf + 0) = 0; // rgvarg
         *(nint*)(dispParamsBuf + 8) = 0; // rgdispidNamedArgs
         *(uint*)(dispParamsBuf + 16) = 0; // cArgs
         *(uint*)(dispParamsBuf + 20) = 0; // cNamedArgs

         // Allocate VARIANT for result
         byte* variantResult = stackalloc byte[VARIANT_SIZE];
         new Span<byte>(variantResult, VARIANT_SIZE).Clear();

         int hr = invokeFn(
            dispatchPtr,
            DISPID_NEWENUM,
            &iidNull,
            0, // LCID
            DISPATCH_PROPERTYGET,
            (nint)dispParamsBuf,
            (nint)variantResult,
            0, // pExcepInfo
            0  // puArgErr
         );

         if (hr != 0)
            return 0;

         ushort vt = *(ushort*)variantResult;
         nint punkEnum = *(nint*)(variantResult + 8);

         if ((vt == VT_UNKNOWN || vt == VT_DISPATCH) && punkEnum != 0)
         {
            // QueryInterface for IEnumVARIANT
            var iidEnum = IID_IEnumVARIANT;
            nint enumPtr;
            nint* enumVtable = *(nint**)punkEnum;
            var queryFn = (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)enumVtable[0];

            if (queryFn(punkEnum, &iidEnum, &enumPtr) == 0)
            {
               ComHelper.Release(punkEnum);
               return enumPtr;
            }

            ComHelper.Release(punkEnum);
         }

         return 0;
      }
   }
}
