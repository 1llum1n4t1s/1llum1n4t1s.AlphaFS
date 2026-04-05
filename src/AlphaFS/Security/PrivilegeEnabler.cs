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
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Alphaleonis.Win32.Security
{
   /// <summary>1つ以上の特権を有効にするために使用されます。指定された特権はインスタンスの存続期間中に有効になります。昇格された特権が不要になったときに適切にDisposeされるように、<c>using</c>ステートメント内でこのオブジェクトのインスタンスを作成してください。</summary>
   public sealed class PrivilegeEnabler : IDisposable
   {
      #region PrivilegeEnabler

      private readonly List<InternalPrivilegeEnabler> _enabledPrivileges = new List<InternalPrivilegeEnabler>();

      /// <summary><see cref="PrivilegeEnabler"/>クラスの新しいインスタンスを初期化します。
      /// これにより、指定された特権が有効になり（既に有効でない場合）、オブジェクトがDisposeされたときに再び無効になります。
      /// （既に有効な特権��無効にされません。）
      /// </summary>
      /// <param name="privilege">有効にする特権。</param>
      /// <param name="privileges">追加で有効にする特権。</param>
      public PrivilegeEnabler(Privilege privilege, params Privilege[] privileges)
      {
         _enabledPrivileges.Add(new InternalPrivilegeEnabler(privilege));

         if (privileges != null)
         {
            foreach (var priv in privileges)
               _enabledPrivileges.Add(new InternalPrivilegeEnabler(priv));
         }
      }

      #endregion // PrivilegeEnabler

      #region Dispose

      /// <summary>このインスタンスによって有効にされた特権が確実に無効にされるようにします。</summary>
      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      public void Dispose()
      {
         foreach (var t in _enabledPrivileges)
         {
            try
            {
               t.Dispose();
            }
            catch
            {
               // ここではすべての例外を無視する
            }
         }
      }

      #endregion // Dispose

      #region EnabledPrivileges

      /// <summary>有効にされた特権を取得します。コンストラクタで指定されたすべての特権が含まれない場合があります。このインスタンスによって実際に有効にされた特権のみが返されます。</summary>
      /// <value>有効にされた特権。</value>
      public IEnumerable<Privilege> EnabledPrivileges
      {
         get { return from priv in _enabledPrivileges where priv.EnabledPrivilege != null select priv.EnabledPrivilege; }
      }

      #endregion // EnabledPrivileges
   }
}