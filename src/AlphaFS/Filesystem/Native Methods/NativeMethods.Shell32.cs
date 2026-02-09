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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      #region AssocXxx

      /// <summary>Returns a pointer to an IQueryAssociations object.</summary>
      /// <returns>If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
      /// <remarks>Minimum supported client: Windows 2000 Professional, Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint AssocCreate(Guid clsid, ref Guid riid, out IntPtr ppv);

      /// <summary>Searches for and retrieves a file or protocol association-related string from the registry.</summary>
      /// <returns>Return value Type: HRESULT. Returns a standard COM error value, including the following: S_OK, E_POINTER and S_FALSE.</returns>
      /// <remarks>Minimum supported client: Windows 2000 Professional</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "AssocQueryStringW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint AssocQueryString(Shell32.AssociationAttributes flags, Shell32.AssociationString str, [MarshalAs(UnmanagedType.LPWStr)] string pszAssoc, [MarshalAs(UnmanagedType.LPWStr)] string pszExtra, StringBuilder pszOut, [MarshalAs(UnmanagedType.U4)] out uint pcchOut);


      #region IQueryAssociations

      internal static readonly Guid ClsidQueryAssociations = new Guid("A07034FD-6CAA-4954-AC3F-97A27216F98A");
      internal const string QueryAssociationsGuid = "C46CA590-3C3F-11D2-BEE6-0000F805CA57";

      /// <summary>AOT-safe wrapper for the IQueryAssociations COM interface (IUnknown-based).
      /// Uses raw vtable function pointer calls instead of runtime COM interop.</summary>
      internal readonly unsafe struct QueryAssociationsWrapper : IDisposable
      {
         private readonly nint _ptr;

         internal QueryAssociationsWrapper(nint comPtr)
         {
            _ptr = comPtr;
         }

         internal bool IsValid => _ptr != 0;

         /// <summary>Initializes the IQueryAssociations interface and sets the root key to the appropriate ProgID.</summary>
         internal void Init(Shell32.AssociationAttributes flags, string pszAssoc, nint hkProgid, nint hwnd)
         {
            // IUnknown vtable: [0] QueryInterface, [1] AddRef, [2] Release
            // IQueryAssociations vtable: [3] Init, [4] GetKey, [5] GetString, [6] GetData
            nint* vtable = *(nint**)_ptr;
            var initFn = (delegate* unmanaged[Stdcall]<nint, Shell32.AssociationAttributes, char*, nint, nint, int>)vtable[3];

            int hr;
            fixed (char* pAssoc = pszAssoc)
            {
               hr = initFn(_ptr, flags, pAssoc, hkProgid, hwnd);
            }

            Marshal.ThrowExceptionForHR(hr);
         }

         /// <summary>Searches for and retrieves a file or protocol association-related string from the registry.</summary>
         internal void GetString(Shell32.AssociationAttributes flags, Shell32.AssociationString str, string pwszExtra, StringBuilder pwszOut, out int pcchOut)
         {
            // IQueryAssociations vtable: [3] Init, [4] GetKey, [5] GetString
            nint* vtable = *(nint**)_ptr;
            var getStringFn = (delegate* unmanaged[Stdcall]<nint, Shell32.AssociationAttributes, Shell32.AssociationString, char*, char*, int*, int>)vtable[5];

            int hr;
            pcchOut = pwszOut.Capacity;
            // Allocate a temporary buffer for the output
            var buffer = new char[pcchOut];
            fixed (char* pExtra = pwszExtra)
            fixed (char* pOut = buffer)
            fixed (int* pSize = &pcchOut)
            {
               hr = getStringFn(_ptr, flags, str, pExtra, pOut, pSize);
            }

            if (hr >= 0)
            {
               pwszOut.Clear();
               pwszOut.Append(buffer, 0, pcchOut > 0 ? pcchOut - 1 : 0);
            }

            Marshal.ThrowExceptionForHR(hr);
         }

         /// <summary>Releases the COM object reference.</summary>
         public void Dispose()
         {
            if (_ptr != 0)
            {
               nint* vtable = *(nint**)_ptr;
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)vtable[2];
               releaseFn(_ptr);
            }
         }
      }

      /// <summary>Creates a new IQueryAssociations wrapper from AssocCreate.</summary>
      internal static QueryAssociationsWrapper CreateQueryAssociations()
      {
         var iid = new Guid(QueryAssociationsGuid);
         var hr = AssocCreate(ClsidQueryAssociations, ref iid, out var ptr);
         return hr == Win32Errors.S_OK ? new QueryAssociationsWrapper(ptr) : new QueryAssociationsWrapper(0);
      }

      #endregion // IQueryAssociations

      #endregion // AssocXxx


      #region Path

      /// <summary>Determines whether a path to a file system object such as a file or folder is valid.</summary>
      /// <returns><c>true</c> if the file exists; otherwise, <c>false</c>. Call GetLastError for extended error information.</returns>
      /// <remarks>
      /// This function tests the validity of the path.
      /// A path specified by Universal Naming Convention (UNC) is limited to a file only; that is, \\server\share\file is permitted.
      /// A network share path to a server or server share is not permitted; that is, \\server or \\server\share.
      /// This function returns FALSE if a mounted remote drive is out of service.
      /// </remarks>
      /// <remarks>Minimum supported client: Windows 2000 Professional</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "PathFileExistsW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool PathFileExists([MarshalAs(UnmanagedType.LPWStr)] string pszPath);


      /// <summary>Converts a file URL to a Microsoft MS-DOS path.</summary>
      /// <returns>Type: HRESULT
      /// If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
      /// </returns>
      /// <remarks>Minimum supported client: Windows 2000 Professional, Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "PathCreateFromUrlW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint PathCreateFromUrl([MarshalAs(UnmanagedType.LPWStr)] string pszUrl, StringBuilder pszPath, [MarshalAs(UnmanagedType.U4)] ref uint pcchPath, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>Creates a path from a file URL.</summary>
      /// <returns>Type: HRESULT
      /// If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
      /// </returns>
      /// <remarks>Minimum supported client: Windows Vista [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows Server 2008 [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint PathCreateFromUrlAlloc([MarshalAs(UnmanagedType.LPWStr)] string pszIn, out StringBuilder pszPath, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>Converts a Microsoft MS-DOS path to a canonicalized URL.</summary>
      /// <returns>Type: HRESULT
      /// Returns S_FALSE if pszPath is already in URL format. In this case, pszPath will simply be copied to pszUrl.
      /// Otherwise, it returns S_OK if successful or a standard COM error value if not.
      /// </returns>
      /// <remarks>
      /// UrlCreateFromPath does not support extended paths. These are paths that include the extended-length path prefix "\\?\".
      /// </remarks>
      /// <remarks>Minimum supported client: Windows 2000 Professional, Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "UrlCreateFromPathW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint UrlCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, StringBuilder pszUrl, ref uint pcchUrl, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>Tests whether a URL is a specified type.</summary>
      /// <returns>
      /// Type: BOOL
      /// For all but one of the URL types, UrlIs returns <c>true</c> if the URL is the specified type, <c>true</c> otherwise.
      /// If UrlIs is set to <see cref="Shell32.UrlType.IsAppliable"/>, UrlIs will attempt to determine the URL scheme.
      /// If the function is able to determine a scheme, it returns <c>true</c>, or <c>false</c>.
      /// </returns>
      /// <remarks>Minimum supported client: Windows 2000 Professional, Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "UrlIsW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool UrlIs([MarshalAs(UnmanagedType.LPWStr)] string pszUrl, Shell32.UrlType urlIs);

      #endregion // Path


      /// <summary>Destroys an icon and frees any memory the icon occupied.</summary>
      /// <remarks>Minimum supported client: Windows XP [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("user32.dll", SetLastError = false)]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DestroyIcon(IntPtr hIcon);


      /// <summary>Retrieves information about an object in the file system, such as a file, folder, directory, or drive root.</summary>
      /// <remarks>You should call this function from a background thread. Failure to do so could cause the UI to stop responding.</remarks>
      /// <remarks>Minimum supported client: Windows 2000 Professional [desktop apps only]</remarks>
      /// <remarks>Minimum supported server: Windows 2000 Server [desktop apps only]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SHGetFileInfoW"), SuppressUnmanagedCodeSecurity]
      internal static extern IntPtr ShGetFileInfo([MarshalAs(UnmanagedType.LPWStr)] string pszPath, FileAttributes dwFileAttributes, [MarshalAs(UnmanagedType.Struct)] out Shell32.FileInfo psfi, [MarshalAs(UnmanagedType.U4)] uint cbFileInfo, [MarshalAs(UnmanagedType.U4)] Shell32.FileAttributes uFlags);
   }
}
