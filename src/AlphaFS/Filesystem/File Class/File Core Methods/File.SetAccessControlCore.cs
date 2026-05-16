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
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using Alphaleonis.Win32.Security;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>[AlphaFS] Applies access control list (ACL) entries described by a <see cref="FileSecurity"/>/<see cref="DirectorySecurity"/> object to the specified file or directory.</summary>
      /// <remarks><paramref name="path"/>または<paramref name="handle"/>のいずれかを使用し、両方は使用しないでください。</remarks>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">アクセス制御リスト(ACL)エントリを追加または削除するファイルまたはディレクトリ。このパラメータは<c>null</c>にできます。</param>
      /// <param name="handle">A <see cref="SafeFileHandle"/> to add or remove access control list (ACL) entries from. This parameter This parameter may be <c>null</c>.</param>
      /// <param name="objectSecurity"><paramref name="path"/>/<paramref name="handle"/>パラメータで記述されたファイルまたはディレクトリに適用するACLエントリを記述する<see cref="FileSecurity"/>/<see cref="DirectorySecurity"/>オブジェクト。</param>
      /// <param name="includeSections">設定するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ以上。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
      [SecurityCritical]
      internal static void SetAccessControlCore(string path, SafeFileHandle handle, ObjectSecurity objectSecurity, AccessControlSections includeSections, PathFormat pathFormat)
      {
         if (pathFormat == PathFormat.RelativePath)
         {
            Path.CheckSupportedPathFormat(path, true, true);
         }

         if (objectSecurity == null)
         {
            throw new ArgumentNullException("objectSecurity");
         }


         var managedDescriptor = objectSecurity.GetSecurityDescriptorBinaryForm();

         using var safeBuffer = new SafeGlobalMemoryBufferHandle(managedDescriptor.Length);
         var pathLp = Path.GetExtendedLengthPathCore(null, path, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.CheckInvalidPathChars);

         safeBuffer.CopyFrom(managedDescriptor, 0, managedDescriptor.Length);

         uint revision;


         var success = Security.NativeMethods.GetSecurityDescriptorControl(safeBuffer, out var control, out revision);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError, pathLp);
         }


         PrivilegeEnabler privilegeEnabler = null;

         try
         {
            var securityInfo = SECURITY_INFORMATION.None;
            var pDacl = IntPtr.Zero;

            if ((includeSections & AccessControlSections.Access) != 0)
            {
               bool daclDefaulted;


               success = Security.NativeMethods.GetSecurityDescriptorDacl(safeBuffer, out var daclPresent, out pDacl, out daclDefaulted);

               lastError = Marshal.GetLastWin32Error();
               if (!success)
               {
                  NativeError.ThrowException(lastError, pathLp);
               }


               if (daclPresent)
               {
                  securityInfo |= SECURITY_INFORMATION.DACL_SECURITY_INFORMATION;
                  securityInfo |= (control & SECURITY_DESCRIPTOR_CONTROL.SE_DACL_PROTECTED) != 0 ? SECURITY_INFORMATION.PROTECTED_DACL_SECURITY_INFORMATION : SECURITY_INFORMATION.UNPROTECTED_DACL_SECURITY_INFORMATION;
               }
            }


            var pSacl = IntPtr.Zero;

            if ((includeSections & AccessControlSections.Audit) != 0)
            {
               bool saclDefaulted;


               success = Security.NativeMethods.GetSecurityDescriptorSacl(safeBuffer, out var saclPresent, out pSacl, out saclDefaulted);

               lastError = Marshal.GetLastWin32Error();
               if (!success)
               {
                  NativeError.ThrowException(lastError, pathLp);
               }


               if (saclPresent)
               {
                  securityInfo |= SECURITY_INFORMATION.SACL_SECURITY_INFORMATION;
                  securityInfo |= (control & SECURITY_DESCRIPTOR_CONTROL.SE_SACL_PROTECTED) != 0 ? SECURITY_INFORMATION.PROTECTED_SACL_SECURITY_INFORMATION : SECURITY_INFORMATION.UNPROTECTED_SACL_SECURITY_INFORMATION;

                  privilegeEnabler = new PrivilegeEnabler(Privilege.Security);
               }
            }


            var pOwner = IntPtr.Zero;

            if ((includeSections & AccessControlSections.Owner) != 0)
            {
               bool ownerDefaulted;


               success = Security.NativeMethods.GetSecurityDescriptorOwner(safeBuffer, out pOwner, out ownerDefaulted);

               lastError = Marshal.GetLastWin32Error();
               if (!success)
               {
                  NativeError.ThrowException(lastError, pathLp);
               }


               if (pOwner != IntPtr.Zero)
               {
                  securityInfo |= SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION;
               }
            }


            var pGroup = IntPtr.Zero;

            if ((includeSections & AccessControlSections.Group) != 0)
            {
               bool groupDefaulted;


               success = Security.NativeMethods.GetSecurityDescriptorGroup(safeBuffer, out pGroup, out groupDefaulted);

               lastError = Marshal.GetLastWin32Error();
               if (!success)
               {
                  NativeError.ThrowException(lastError, pathLp);
               }


               if (pGroup != IntPtr.Zero)
               {
                  securityInfo |= SECURITY_INFORMATION.GROUP_SECURITY_INFORMATION;
               }
            }




            if (!Utils.IsNullOrWhiteSpace(pathLp))
            {
               // SetNamedSecurityInfo()
               // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。

               lastError = (int) Security.NativeMethods.SetNamedSecurityInfo(pathLp, SE_OBJECT_TYPE.SE_FILE_OBJECT, securityInfo, pOwner, pGroup, pDacl, pSacl);

               if (lastError != Win32Errors.ERROR_SUCCESS)
               {
                  NativeError.ThrowException(lastError, pathLp);
               }
            }

            else
            {
               if (NativeMethods.IsValidHandle(handle))
               {
                  lastError = (int) Security.NativeMethods.SetSecurityInfo(handle, SE_OBJECT_TYPE.SE_FILE_OBJECT, securityInfo, pOwner, pGroup, pDacl, pSacl);

                  if (lastError != Win32Errors.ERROR_SUCCESS)
                  {
                     NativeError.ThrowException(lastError);
                  }
               }
            }
         }
         finally
         {
            if (null != privilegeEnabler)
            {
               privilegeEnabler.Dispose();
            }
         }
      }
   }
}
