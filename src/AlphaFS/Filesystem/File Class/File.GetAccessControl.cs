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
using System.IO;
using System.Security;
using System.Security.AccessControl;
using Alphaleonis.Win32.Security;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      /// <summary>Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file.</summary>
      /// <returns><paramref name="path"/>パラメータで記述されたファイルのアクセス制御ルールをカプセル化する<see cref="FileSecurity"/>オブジェクト。</returns>      
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">ファイルのアクセス制御リスト(ACL)情報を記述する<see cref="FileSecurity"/>オブジェクトを含むファイルのパス。</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(string path)
      {
         return GetAccessControlCore<FileSecurity>(false, path, AccessControlSections.Access | AccessControlSections.Group | AccessControlSections.Owner, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file.</summary>
      /// <returns><paramref name="path"/>パラメータで記述されたファイルのアクセス制御ルールをカプセル化する<see cref="FileSecurity"/>オブジェクト。</returns>      
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">ファイルのアクセス制御リスト(ACL)情報を記述する<see cref="FileSecurity"/>オブジェクトを含むファイルのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(string path, PathFormat pathFormat)
      {
         return GetAccessControlCore<FileSecurity>(false, path, AccessControlSections.Access | AccessControlSections.Group | AccessControlSections.Owner, pathFormat);
      }


      /// <summary>Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file.</summary>
      /// <returns><paramref name="path"/>パラメータで記述されたファイルのアクセス制御ルールをカプセル化する<see cref="FileSecurity"/>オブジェクト。</returns>      
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">ファイルのアクセス制御リスト(ACL)情報を記述する<see cref="FileSecurity"/>オブジェクトを含むファイルのパス。</param>
      /// <param name="includeSections">受信するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ(または複数)。</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(string path, AccessControlSections includeSections)
      {
         return GetAccessControlCore<FileSecurity>(false, path, includeSections, PathFormat.RelativePath);
      }
      

      /// <summary>[AlphaFS] Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file.</summary>
      /// <returns><paramref name="path"/>パラメータで記述されたファイルのアクセス制御ルールをカプセル化する<see cref="FileSecurity"/>オブジェクト。</returns>
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="path">ファイルのアクセス制御リスト(ACL)情報を記述する<see cref="FileSecurity"/>オブジェクトを含むファイルのパス。</param>
      /// <param name="includeSections">受信するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ(または複数)。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(string path, AccessControlSections includeSections, PathFormat pathFormat)
      {
         return GetAccessControlCore<FileSecurity>(false, path, includeSections, pathFormat);
      }


      /// <summary>[AlphaFS] Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file handle.</summary>
      /// <returns>A <see cref="FileSecurity"/> object that encapsulates the access control rules for the file described by the <paramref name="handle"/> parameter.</returns>      
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="handle">A <see cref="SafeFileHandle"/> to a file containing a <see cref="FileSecurity"/> object that describes the file's access control list (ACL) information.</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(SafeFileHandle handle)
      {
         return GetAccessControlHandleCore<FileSecurity>(false, false, handle, AccessControlSections.Access | AccessControlSections.Group | AccessControlSections.Owner, SECURITY_INFORMATION.None);
      }


      /// <summary>[AlphaFS] Gets a <see cref="FileSecurity"/> object that encapsulates the access control list (ACL) entries for a specified file handle.</summary>
      /// <returns>A <see cref="FileSecurity"/> object that encapsulates the access control rules for the file described by the <paramref name="handle"/> parameter.</returns>      
      /// <exception cref="IOException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="ArgumentNullException"/>
      /// <param name="handle">A <see cref="SafeFileHandle"/> to a file containing a <see cref="FileSecurity"/> object that describes the file's access control list (ACL) information.</param>
      /// <param name="includeSections">受信するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ(または複数)。</param>
      [SecurityCritical]
      public static FileSecurity GetAccessControl(SafeFileHandle handle, AccessControlSections includeSections)
      {
         return GetAccessControlHandleCore<FileSecurity>(false, false, handle, includeSections, SECURITY_INFORMATION.None);
      }
   }
}
