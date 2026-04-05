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
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>現在の DirectoryInfo オブジェクトで記述されたディレクトリのアクセス制御リスト (ACL) エントリをカプセル化する <see cref="DirectorySecurity"/> オブジェクトを取得します。</summary>
      /// <returns>ディレクトリのアクセス制御規則をカプセル化する <see cref="DirectorySecurity"/> オブジェクト。</returns>
      [SecurityCritical]
      public DirectorySecurity GetAccessControl()
      {
         return File.GetAccessControlCore<DirectorySecurity>(true, LongFullName, AccessControlSections.Access | AccessControlSections.Group | AccessControlSections.Owner, PathFormat.LongFullPath);
      }


      /// <summary>現在の <see cref="DirectoryInfo"/> オブジェクトで記述されたディレクトリの指定された種類のアクセス制御リスト (ACL) エントリをカプセル化する <see cref="DirectorySecurity"/> オブジェクトを取得します。</summary>
      /// <param name="includeSections">受信するアクセス制御リスト (ACL) 情報の種類を指定する <see cref="AccessControlSections"/> 値の 1 つ。</param>
      /// <returns>path パラメーターで記述されたファイルのアクセス制御規則をカプセル化する <see cref="DirectorySecurity"/> オブジェクト。</returns>
      [SecurityCritical]
      public DirectorySecurity GetAccessControl(AccessControlSections includeSections)
      {
         return File.GetAccessControlCore<DirectorySecurity>(true, LongFullName, includeSections, PathFormat.LongFullPath);
      }

      #endregion // .NET
   }
}
