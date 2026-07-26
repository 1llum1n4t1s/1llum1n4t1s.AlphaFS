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

using System.Diagnostics.CodeAnalysis;
using System.Security;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Filesystem
{
   public sealed partial class DirectoryInfo
   {
      #region .NET

      /// <summary>現在の DirectoryInfo オブジェクトで記述されたディレクトリに、<see cref="DirectorySecurity"/> オブジェクトで記述されたアクセス制御リスト (ACL) エントリを適用します。</summary>
      /// <param name="directorySecurity">path パラメーターで記述されたディレクトリに適用する ACL エントリを記述する <see cref="DirectorySecurity"/> オブジェクト。</param>
      /// <remarks>
      ///   既定では DACL (<see cref="AccessControlSections.Access"/>) だけを適用します。
      ///   所有者・グループ・監査 (SACL) も書き込むには <c>includeSections</c> を取るオーバーロードを使ってください。
      ///   <para>所有者とグループの書き込みには対象に対する WRITE_OWNER が、SACL には SeSecurityPrivilege が必要です。
      ///   GetAccessControl → ルール追加 → SetAccessControl という通常の流れでは DACL しか変更していないので、
      ///   それ以外を既定で書きに行くと、権限のない環境で不要に (5) Access is denied で失敗します。</para>
      /// </remarks>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public void SetAccessControl(DirectorySecurity directorySecurity)
      {
         File.SetAccessControlCore(LongFullName, null, directorySecurity, AccessControlSections.Access, PathFormat.LongFullPath);
      }


      /// <summary>現在の DirectoryInfo オブジェクトで記述されたディレクトリに、<see cref="DirectorySecurity"/> オブジェクトで記述されたアクセス制御リスト (ACL) エントリを適用します。</summary>
      /// <param name="directorySecurity">path パラメーターで記述されたディレクトリに適用する ACL エントリを記述する <see cref="DirectorySecurity"/> オブジェクト。</param>
      /// <param name="includeSections">設定するアクセス制御リスト (ACL) 情報の種類を指定する <see cref="AccessControlSections"/> 値の 1 つ以上。</param>
      [SuppressMessage("Microsoft.Design", "CA1011:ConsiderPassingBaseTypesAsParameters")]
      [SecurityCritical]
      public void SetAccessControl(DirectorySecurity directorySecurity, AccessControlSections includeSections)
      {
         File.SetAccessControlCore(LongFullName, null, directorySecurity, includeSections, PathFormat.LongFullPath);
      }

      #endregion // .NET
   }
}
