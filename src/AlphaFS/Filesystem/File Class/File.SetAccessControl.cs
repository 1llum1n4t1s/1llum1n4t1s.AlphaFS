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
using System.Security;
using System.Security.AccessControl;
using Microsoft.Win32.SafeHandles;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region .NET

      /// <summary>Applies access control list (ACL) entries described by a <see cref="FileSecurity"/> FileSecurity object to the specified file.</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">A file to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity"><paramref name="path"/>パラメータで記述されたファイルに適用するACLエントリを記述する<see cref="FileSecurity"/>オブジェクト。</param>      
      /// <remarks>
      ///   既定では DACL (<see cref="AccessControlSections.Access"/>) だけを適用します。
      ///   所有者・グループ・監査 (SACL) も書き込むには <c>includeSections</c> を取るオーバーロードを使ってください。
      ///   <para>所有者とグループの書き込みには対象に対する WRITE_OWNER が、SACL には SeSecurityPrivilege が必要です。
      ///   GetAccessControl → ルール追加 → SetAccessControl という通常の流れでは DACL しか変更していないので、
      ///   それ以外を既定で書きに行くと、権限のない環境で不要に (5) Access is denied で失敗します。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(string path, FileSecurity fileSecurity)
      {
         SetAccessControlCore(path, null, fileSecurity, AccessControlSections.Access, PathFormat.RelativePath);
      }

      #endregion // .NET


      /// <summary><see cref="DirectorySecurity"/>オブジェクトで記述されたアクセス制御リスト(ACL)エントリを指定されたディレクトリに適用します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">A directory to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity">A <see cref="FileSecurity "/> object that describes an ACL entry to apply to the directory described by the path parameter.</param>
      /// <param name="includeSections">設定するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ以上。</param>      
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(string path, FileSecurity fileSecurity, AccessControlSections includeSections)
      {
         SetAccessControlCore(path, null, fileSecurity, includeSections, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] <see cref="FileSecurity"/> FileSecurityオブジェクトで記述されたアクセス制御リスト(ACL)エントリを指定されたファイルに適用します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">A file to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity"><paramref name="path"/>パラメータで記述されたファイルに適用するACLエントリを記述する<see cref="FileSecurity"/>オブジェクト。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      /// <remarks>
      ///   既定では DACL (<see cref="AccessControlSections.Access"/>) だけを適用します。
      ///   所有者・グループ・監査 (SACL) も書き込むには <c>includeSections</c> を取るオーバーロードを使ってください。
      ///   <para>所有者とグループの書き込みには対象に対する WRITE_OWNER が、SACL には SeSecurityPrivilege が必要です。
      ///   GetAccessControl → ルール追加 → SetAccessControl という通常の流れでは DACL しか変更していないので、
      ///   それ以外を既定で書きに行くと、権限のない環境で不要に (5) Access is denied で失敗します。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(string path, FileSecurity fileSecurity, PathFormat pathFormat)
      {
         SetAccessControlCore(path, null, fileSecurity, AccessControlSections.Access, pathFormat);
      }


      /// <summary>[AlphaFS] <see cref="DirectorySecurity"/>オブジェクトで記述されたアクセス制御リスト(ACL)エントリを指定されたディレクトリに適用します。</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="path">A directory to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity">A <see cref="FileSecurity "/> object that describes an ACL entry to apply to the directory described by the path parameter.</param>
      /// <param name="includeSections">設定するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ以上。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(string path, FileSecurity fileSecurity, AccessControlSections includeSections, PathFormat pathFormat)
      {
         SetAccessControlCore(path, null, fileSecurity, includeSections, pathFormat);
      }


      /// <summary>Applies access control list (ACL) entries described by a <see cref="FileSecurity"/> FileSecurity object to the specified file.</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="handle">A <see cref="SafeFileHandle"/> to a file to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity">A <see cref="FileSecurity"/> object that describes an ACL entry to apply to the file described by the <paramref name="handle"/> parameter.</param>      
      /// <remarks>
      ///   既定では DACL (<see cref="AccessControlSections.Access"/>) だけを適用します。
      ///   所有者・グループ・監査 (SACL) も書き込むには <c>includeSections</c> を取るオーバーロードを使ってください。
      ///   <para>所有者とグループの書き込みには対象に対する WRITE_OWNER が、SACL には SeSecurityPrivilege が必要です。
      ///   GetAccessControl → ルール追加 → SetAccessControl という通常の流れでは DACL しか変更していないので、
      ///   それ以外を既定で書きに行くと、権限のない環境で不要に (5) Access is denied で失敗します。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(SafeFileHandle handle, FileSecurity fileSecurity)
      {
         SetAccessControlCore(null, handle, fileSecurity, AccessControlSections.Access, PathFormat.LongFullPath);
      }


      /// <summary>Applies access control list (ACL) entries described by a <see cref="FileSecurity"/> FileSecurity object to the specified file.</summary>
      /// <exception cref="ArgumentNullException"/>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="handle">A <see cref="SafeFileHandle"/> to a file to add or remove access control list (ACL) entries from.</param>
      /// <param name="fileSecurity">A <see cref="FileSecurity"/> object that describes an ACL entry to apply to the file described by the <paramref name="handle"/> parameter.</param>      
      /// <param name="includeSections">設定するアクセス制御リスト(ACL)情報の種類を指定する<see cref="AccessControlSections"/>値の1つ以上。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public static void SetAccessControl(SafeFileHandle handle, FileSecurity fileSecurity, AccessControlSections includeSections)
      {
         SetAccessControlCore(null, handle, fileSecurity, includeSections, PathFormat.LongFullPath);
      }
   }
}
