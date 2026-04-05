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

using Alphaleonis.Win32.Filesystem;
using System;
using System.Globalization;

namespace Alphaleonis.Win32.Network
{
   /// <summary>サーバーメッセージブロック (SMB) 共有に関する情報を含みます。このクラスは継承できません。</summary>
   [Serializable]
   public sealed class ShareInfo
   {
      #region コンストラクター

      /// <summary><see cref="ShareInfo"/> インスタンスを作成します。</summary>
      /// <param name="host">共有を取得するホスト。</param>
      /// <param name="shareLevel"><see cref="ShareInfoLevel"/> オプションのいずれか。</param>
      /// <param name="shareInfo"><see cref="NativeMethods.SHARE_INFO_2"/> または <see cref="NativeMethods.SHARE_INFO_503"/> インスタンス。</param>
      internal ShareInfo(string host, ShareInfoLevel shareLevel, object shareInfo)
      {
         host = host ?? Environment.MachineName;


         switch (shareLevel)
         {
            case ShareInfoLevel.Info1005:
               var s1005 = (NativeMethods.SHARE_INFO_1005) shareInfo;
               ServerName = host;
               ResourceType = s1005.shi1005_flags;
               break;


            case ShareInfoLevel.Info503:
               var s503 = (NativeMethods.SHARE_INFO_503) shareInfo;
               CurrentUses = s503.shi503_current_uses;
               MaxUses = s503.shi503_max_uses;
               NetName = s503.shi503_netname;
               Password = s503.shi503_passwd;
               Path = Utils.IsNullOrWhiteSpace(s503.shi503_path) ? null : s503.shi503_path;
               Permissions = s503.shi503_permissions;
               Remark = s503.shi503_remark;
               ServerName = s503.shi503_servername == "*" ? host : s503.shi503_servername;
               ShareType = s503.shi503_type;
               SecurityDescriptor = s503.shi503_security_descriptor;
               break;


            case ShareInfoLevel.Info502:
               var s502 = (NativeMethods.SHARE_INFO_502) shareInfo;
               CurrentUses = s502.shi502_current_uses;
               MaxUses = s502.shi502_max_uses;
               NetName = s502.shi502_netname;
               Password = s502.shi502_passwd;
               Path = Utils.IsNullOrWhiteSpace(s502.shi502_path) ? null : s502.shi502_path;
               Permissions = s502.shi502_permissions;
               Remark = s502.shi502_remark;
               ServerName = host;
               ShareType = s502.shi502_type;
               SecurityDescriptor = s502.shi502_security_descriptor;
               break;


            case ShareInfoLevel.Info2:
               var s2 = (NativeMethods.SHARE_INFO_2) shareInfo;
               CurrentUses = s2.shi2_current_uses;
               MaxUses = s2.shi2_max_uses;
               NetName = s2.shi2_netname;
               Password = s2.shi2_passwd;
               Path = Utils.IsNullOrWhiteSpace(s2.shi2_path) ? null : s2.shi2_path;
               Permissions = s2.shi2_permissions;
               Remark = s2.shi2_remark;
               ServerName = host;
               ShareType = s2.shi2_type;
               break;


            case ShareInfoLevel.Info1:
               var s1 = (NativeMethods.SHARE_INFO_1) shareInfo;
               CurrentUses = 0;
               MaxUses = 0;
               NetName = s1.shi1_netname;
               Password = null;
               Path = null;
               Permissions = AccessPermissions.None;
               Remark = s1.shi1_remark;
               ServerName = host;
               ShareType = s1.shi1_type;
               break;
         }


         NetFullPath = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}{3}", Filesystem.Path.UncPrefix, ServerName, Filesystem.Path.DirectorySeparatorChar, NetName);

         ShareLevel = shareLevel;
      }

      #endregion // コンストラクター


      #region メソッド

      /// <summary>共有へのフルパスを返します。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return NetFullPath;
      }

      #endregion // メソッド


      #region プロパティ

      /// <summary>リソースへの現在の接続数。</summary>
      public long CurrentUses { get; private set; }


      private DirectoryInfo _directoryInfo;

      /// <summary>この共有に関連付けられた <see cref="DirectoryInfo"/> インスタンス。</summary>
      public DirectoryInfo DirectoryInfo
      {
         get { return _directoryInfo ?? (_directoryInfo = new DirectoryInfo(null, NetFullPath, PathFormat.FullPath)); }
      }


      /// <summary>共有への完全な UNC パスを返します。</summary>
      public string NetFullPath { get; internal set; }


      /// <summary>共有リソースが収容できる同時接続の最大数。</summary>
      /// <remarks>このメンバーで指定された値が -1 の場合、接続数は無制限です。</remarks>
      public long MaxUses { get; private set; }


      /// <summary>共有リソースの名前。</summary>
      public string NetName { get; private set; }


      /// <summary>共有のパスワード（サーバーが共有レベルのセキュリティで実行されている場合）。</summary>
      public string Password { get; private set; }


      /// <summary>共有リソースのローカルパス。</summary>
      /// <remarks>ディスクの場合、このメンバーは共有されているパスです。印刷キューの場合、このメンバーは共有されている印刷キューの名前です。</remarks>
      public string Path { get; private set; }


      /// <summary>共有レベルのセキュリティで実行されているサーバーの共有リソースのアクセス許可。</summary>
      /// <remarks>Windows は共有レベルのセキュリティをサポートしていないことに注意してください。このメンバーはユーザーレベルのセキュリティで実行されているサーバーでは無視されます。</remarks>
      public AccessPermissions Permissions { get; private set; }


      /// <summary>共有リソースに関するオプションのコメント。</summary>
      public string Remark { get; private set; }


      /// <summary>この共有に関連付けられた SECURITY_DESCRIPTOR を指定します。</summary>
      public IntPtr SecurityDescriptor { get; private set; }


      /// <summary>共有リソースが存在するリモートサーバーの DNS 名または NetBIOS 名を指定する文字列へのポインター。</summary>
      /// <remarks>"*" の値は、構成されたサーバー名がないことを示します。</remarks>
      public string ServerName { get; private set; }


      /// <summary>共有の種類。</summary>
      public ShareType ShareType { get; private set; }


      private ShareResourceTypes _shareResourceType;

      /// <summary>共有リソースの種類。</summary>
      public ShareResourceTypes ResourceType
      {
         get
         {
            if (_shareResourceType == ShareResourceTypes.None && !Utils.IsNullOrWhiteSpace(NetName))
            {
               _shareResourceType = Host.GetShareInfoCore(ShareInfoLevel.Info1005, ServerName, NetName, true).ResourceType;
            }

            return _shareResourceType;
         }

         private set { _shareResourceType = value; }
      }


      /// <summary><see cref="ShareInfo"/> インスタンスの構造体レベル。</summary>
      public ShareInfoLevel ShareLevel { get; private set; }

      #endregion // プロパティ
   }
}
