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
using System.Globalization;
using Alphaleonis.Win32.Filesystem;

namespace Alphaleonis.Win32.Network
{
   /// <summary>ファイル、デバイス、パイプの識別番号およびその他の関連情報を含みます。このクラスは継承できません。</summary>
   [Serializable]
   public sealed class OpenResourceInfo
   {
      #region コンストラクター

      /// <summary>OpenResourceInfo インスタンスを作成します。</summary>
      internal OpenResourceInfo(string hostName, NativeMethods.FILE_INFO_3 fileInfo)
      {
         HostName = hostName;
         Id = fileInfo.fi3_id;
         Permissions = fileInfo.fi3_permissions;
         TotalLocks = fileInfo.fi3_num_locks;
         PathName = fileInfo.fi3_pathname.Replace(Path.UncPrefix, Path.DirectorySeparator);
         UserName = fileInfo.fi3_username;
      }

      #endregion // コンストラクター


      #region メソッド

      /// <summary>開いているリソースを強制的に閉じます。</summary>
      /// <remarks>このメソッドはファイルを閉じる前にクライアントシステムにキャッシュされたデータをファイルに書き込まないため、注意して使用する必要があります。</remarks>
      public void Close()
      {
         var lastError = NativeMethods.NetFileClose(HostName, (uint) Id);

         if (lastError != Win32Errors.NERR_Success && lastError != Win32Errors.NERR_FileIdNotFound)

         {
            NativeError.ThrowException(lastError, HostName, PathName);
         }
      }

      /// <summary>共有へのフルパスを返します。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return Id.ToString(CultureInfo.InvariantCulture);
      }


      #endregion // メソッド

      #region プロパティ

      /// <summary>ローカルまたはリモートホスト。</summary>
      [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
      [Obsolete("Use HostName")]
      public string Host { get; private set; }

      /// <summary>このリソース情報のホスト名。</summary>
      public string HostName { get; private set; }

      /// <summary>リソースが開かれたときに割り当てられる識別番号。</summary>
      public long Id { get; private set; }

      /// <summary>開かれたリソースのパス。</summary>
      public string PathName { get; private set; }

      /// <summary>オープンしたアプリケーションに関連付けられたアクセス許可。このメンバーは以下の <see cref="AccessPermissions"/> 値の 1 つ以上を指定できます。</summary>
      public AccessPermissions Permissions { get; private set; }

      /// <summary>ファイル、デバイス、またはパイプのファイルロック数。</summary>
      public long TotalLocks { get; private set; }

      /// <summary>リソースを開いたユーザー（ユーザーレベルのセキュリティを持つサーバー上）またはコンピューター（共有レベルのセキュリティを持つサーバー上）を指定します。</summary>
      public string UserName { get; private set; }

      #endregion // プロパティ
   }
}
