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

using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   partial class FileInfo
   {
      #region .NET

      /// <summary>現在の <see cref="FileInfo"/> オブジェクトで記述されたファイルのアクセス制御リスト (ACL) エントリをカプセル化する <see cref="FileSecurity"/> オブジェクトを取得します。</summary>
      /// <returns>現在のファイルのアクセス制御規則をカプセル化する <see cref="FileSecurity"/> オブジェクト。</returns>
      [SecurityCritical]
      public FileSecurity GetAccessControl()
      {
         return File.GetAccessControlCore<FileSecurity>(false, LongFullName, AccessControlSections.Access | AccessControlSections.Group | AccessControlSections.Owner, PathFormat.LongFullPath);
      }


      /// <summary>現在の FileInfo オブジェクトで記述されたファイルの指定された種類のアクセス制御リスト (ACL) エントリをカプセル化する <see cref="FileSecurity"/> オブジェクトを取得します。</summary>
      /// <returns>現在の FileInfo オブジェクトで記述されたファイルの指定された種類のアクセス制御リスト (ACL) エントリをカプセル化する <see cref="FileSecurity"/> オブジェクト。</returns>
      /// <param name="includeSections">取得するアクセス制御エントリのグループを指定する <see cref="System.Security"/> 値の 1 つ。</param>
      [SecurityCritical]
      public FileSecurity GetAccessControl(AccessControlSections includeSections)
      {
         return File.GetAccessControlCore<FileSecurity>(false, LongFullName, includeSections, PathFormat.LongFullPath);
      }

      #endregion // .NET
   }
}
