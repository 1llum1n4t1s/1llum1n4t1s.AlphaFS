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
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Security
{
   /// <summary>アクセストークンの特権を表します。ローカルマシンで利用可能な特権は、このクラスの静的インスタンスとして利用できます。
   /// 別のシステムの特権を表す<see cref="Privilege"/>を作成するには、これらの静的インスタンスの1つと共にシステム名を指定するコンストラクタを使用します。
   /// </summary>
   /// <seealso cref="PrivilegeEnabler"/>
   [ImmutableObject(true)]
   public class Privilege : IEquatable<Privilege>
   {
      #region System Privileges

      /// <summary>プロセスのプライマリトークンの割り当てに必要です。ユーザー権利: プロセスレベルのトークンの置換。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege AssignPrimaryToken = new Privilege("SeAssignPrimaryTokenPrivilege");


      /// <summary>監査ログエントリの生成に必要です。この特権をセキュアサーバーに付与します。ユーザー権利: セキュリティ監査の生成。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Audit = new Privilege("SeAuditPrivilege");


      /// <summary>バックアップ操作の実行に必要です。この特権により、ファイルに指定されたACLに関係なく、すべてのファイルへの読み取りアクセス制御が付与されます。読み取り以外のアクセス要求は引き続きACLで評価されます。ユーザー権利: ファイルとディレクトリのバックアップ。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Backup = new Privilege("SeBackupPrivilege");


      /// <summary>ファイルまたはディレクトリの変更通知の受信に必要です。この特権により、すべてのトラバーサルアクセスチェックもスキップされます。デフォルトですべてのユーザーに有効です。ユーザー権利: トラバースチェックのバイパス。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege ChangeNotify = new Privilege("SeChangeNotifyPrivilege");


      /// <summary>ターミナルサービスセッション中にグローバル名前空間で名前付きファイルマッピングオブジェクトの作成に必要です。この特権は管理者、サービス、およびローカルシステムアカウントにデフォルトで有効です。ユーザー権利: グローバルオブジェクトの作成。</summary>
      /// <remarks>Windows XP/2000: この特権はサポートされていません。この値はWindows Server 2003、Windows XP SP2、およびWindows 2000 SP4以降でサポートされています。</remarks>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege CreateGlobal = new Privilege("SeCreateGlobalPrivilege");


      /// <summary>ページングファイルの作成に必要です。ユーザー権利: ページファイルの作成。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Pagefile")]
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege CreatePagefile = new Privilege("SeCreatePagefilePrivilege");


      /// <summary>永続オブジェクトの作成に必要です。ユーザー権利: 永続共有オブジェクトの作成。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege CreatePermanent = new Privilege("SeCreatePermanentPrivilege");


      /// <summary>シンボリックリンクの作成に必要です。ユーザー権利: シンボリックリンクの作成。</summary>           
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege CreateSymbolicLink = new Privilege("SeCreateSymbolicLinkPrivilege");


      /// <summary>プライマリトークンの作成に必要です。ユーザー権利: トークンオブジェクトの作成。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege CreateToken = new Privilege("SeCreateTokenPrivilege");


      /// <summary>別のアカウントが所有するプロセスのメモリのデバッグと調整に必要です。ユーザー権利: プログラムのデバッグ。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Debug = new Privilege("SeDebugPrivilege");


      /// <summary>ユーザーアカウントとコンピューターアカウントを委任に対して信頼済みとしてマークするために必要です。ユーザー権利: コンピューターとユーザーアカウントの委任の信頼を有効にする。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege EnableDelegation = new Privilege("SeEnableDelegationPrivilege");


      /// <summary>偽装に必要です。ユーザー権利: 認証後にクライアントを偽装する。</summary>
      /// <remarks>Windows XP/2000: この特権はサポートされていません。この値はWindows Server 2003、Windows XP SP2、およびWindows 2000 SP4以降でサポートされています。</remarks>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Impersonate = new Privilege("SeImpersonatePrivilege");


      /// <summary>プロセスの基本優先度の引き上げに必要です。ユーザー権利: スケジュールの優先順位の引き上げ。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege IncreaseBasePriority = new Privilege("SeIncreaseBasePriorityPrivilege");


      /// <summary>プロセスに割り当てられたクォータの引き上げに必要です。ユーザー権利: プロセスのメモリクォータの調整。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege IncreaseQuota = new Privilege("SeIncreaseQuotaPrivilege");


      /// <summary>ユーザーのコンテキストで実行されるアプリケーションにより多くのメモリを割り当てるために必要です。ユーザー権利: プロセスワーキングセットの引き上げ。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege IncreaseWorkingSet = new Privilege("SeIncreaseWorkingSetPrivilege");


      /// <summary>デバイスドライバーの読み込みまたはアンロードに必要です。ユーザー権利: デバイスドライバーの読み込みとアンロード。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege LoadDriver = new Privilege("SeLoadDriverPrivilege");


      /// <summary>メモリ内の物理ページのロックに必要です。ユーザー権利: メモリ内のページのロック。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege LockMemory = new Privilege("SeLockMemoryPrivilege");


      /// <summary>コンピューターアカウントの作成に必要です。ユーザー権利: ドメインにワークステーションを追加。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege MachineAccount = new Privilege("SeMachineAccountPrivilege");


      /// <summary>ボリューム管理特権の有効化に必要です。ユーザー権利: ボリューム上のファイルの管理。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege ManageVolume = new Privilege("SeManageVolumePrivilege");


      /// <summary>単一プロセスのプロファイリング情報の収集に必要です。ユーザー権利: 単一プロセスのプロファイル。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege ProfileSingleProcess = new Privilege("SeProfileSingleProcessPrivilege");


      /// <summary>オブジェクトの必須整合性レベルの変更に必要です。ユーザー権利: オブジェクトラベルの変更。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Relabel")]
      public static readonly Privilege Relabel = new Privilege("SeRelabelPrivilege");


      /// <summary>ネットワーク要求を使用したシステムのシャットダウンに必要です。ユーザー権利: リモートシステムからの強制シャットダウン。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege RemoteShutdown = new Privilege("SeRemoteShutdownPrivilege");


      /// <summary>復元操作の実行に必要です。この特権により、ファイルに指定されたACLに関係なく、すべてのファイルへの書き込みアクセス制御が付与されます。書き込み以外のアクセス要求は引き続きACLで評価されます。さらに、この特権により、任意の有効なユーザーまたはグループSIDをファイルの所有者として設定できます。ユーザー権利: ファイルとディレクトリの復元。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Restore = new Privilege("SeRestorePrivilege");


      /// <summary>監査メッセージの制御と表示など、セキュリティ関連の機能の実行に必要です。この特権は保持者をセキュリティオペレーターとして識別します。ユーザー権利: 監査とセキュリティログの管理。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Security = new Privilege("SeSecurityPrivilege");


      /// <summary>ローカルシステムのシャットダウンに必要です。ユーザー権利: システムのシャットダウン。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Shutdown = new Privilege("SeShutdownPrivilege");


      /// <summary>ドメインコントローラーがLDAPディレクトリ同期サービスを使用するために必要です。この特権により、オブジェクトとプロパティの保護に関係なく、ディレクトリ内のすべてのオブジェクトとプロパティの読み取りが可能になります。デフォルトでは、ドメインコントローラーのAdministratorおよびLocalSystemアカウントに割り当てられます。ユーザー権利: ディレクトリサービスデータの同期。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege SyncAgent = new Privilege("SeSyncAgentPrivilege");


      /// <summary>構成情報を格納するためにこのタイプのメモリを使用するシステムの不揮発性RAMの変更に必要です。ユーザー権利: ファームウェア環境値の変更。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege SystemEnvironment = new Privilege("SeSystemEnvironmentPrivilege");


      /// <summary>システム全体のプロファイリング情報の収集に必要です。ユーザー権利: システムパフォーマンスのプロファイル。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege SystemProfile = new Privilege("SeSystemProfilePrivilege");


      /// <summary>システム時刻の変更に必要です。ユーザー権利: システム時刻の変更。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege SystemTime = new Privilege("SeSystemtimePrivilege");


      /// <summary>随意アクセスを付与されずにオブジェクトの所有権を取得するために必要です。この特権により、所有者の値は保持者が正当にオブジェクトの所有者として割り当てることができる値にのみ設定できます。ユーザー権利: ファイルまたはその他のオブジェクトの所有権の取得。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege TakeOwnership = new Privilege("SeTakeOwnershipPrivilege");


      /// <summary>この特権は保持者を信頼されたコンピューターベースの一部として識別します。一部の信頼された保護サブシステムにこの特権が付与されます。ユーザー権利: オペレーティングシステムの一部として動作。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Tcb")]
      public static readonly Privilege Tcb = new Privilege("SeTcbPrivilege");


      /// <summary>コンピューターの内部時計に関連付けられたタイムゾーンの調整に必要です。ユーザー権利: タイムゾーンの変更。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege TimeZone = new Privilege("SeTimeZonePrivilege");


      /// <summary>信頼された呼び出し元として資格情報マネージャーへのアクセスに必要です。ユーザー権利: 信頼された呼び出し元として資格情報マネージャーにアクセス。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Cred")]
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege TrustedCredManAccess = new Privilege("SeTrustedCredManAccessPrivilege");


      /// <summary>ラップトップのドッキング解除に必要です。ユーザー権利: ドッキングステーションからコンピューターを取り外す。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege Undock = new Privilege("SeUndockPrivilege");


      /// <summary>ターミナルデバイスからの未要請入力の読み取りに必要です。ユーザー権利: 該当なし。</summary>
      [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")]
      public static readonly Privilege UnsolicitedInput = new Privilege("SeUnsolicitedInputPrivilege");

      #endregion // System Privileges

      
      #region Fields

      private readonly string _name;
      private readonly string _systemName;

      #endregion // Fields


      #region Constructors

      /// <summary>指定されたシステム上の指定された特権を表す新しい<see cref="Privilege"/>インスタンスを作成します。</summary>
      /// <param name="systemName">システムの名前。</param>
      /// <param name="privilege">特権名のコピー元となる特権。</param>
      public Privilege(string systemName, Privilege privilege)
      {
         if (Utils.IsNullOrWhiteSpace(systemName))
         {
            throw new ArgumentNullException("systemName", Resources.Privilege_Name_Cannot_Be_Empty);
         }

         _systemName = systemName;

         if (null != privilege)
         {
            _name = privilege._name;
         }
      }


      /// <summary>ローカルシステム上の指定された名前の特権を表す新しい<see cref="Privilege"/>インスタンスを作成します。</summary>
      /// <param name="name">特権の名前。</param>
      private Privilege(string name)
      {
         if (Utils.IsNullOrWhiteSpace(name))
         {
            throw new ArgumentNullException("name", Resources.Privilege_Name_Cannot_Be_Empty);
         }

         _name = name;
      }

      #endregion // Constructors
      

      #region Properties

      /// <summary>この特権を識別するシステム名を取得します。</summary>
      /// <value>この特権を識別するシステム名。</value>
      public string Name
      {
         get { return _name; }
      }

      #endregion // Properties

      
      #region Methods

      /// <summary>この特権を表す表示名を取得します。</summary>
      /// <returns>この特権を表す表示名。</returns>
      [SecurityCritical]
      public string LookupDisplayName()
      {
         const uint initialCapacity = 10;
         var bufferSize = initialCapacity;
         var displayName = new StringBuilder((int) bufferSize);
         uint languageId;

      Retry:

         var success = NativeMethods.LookupPrivilegeDisplayName(_systemName, _name, ref displayName, ref bufferSize, out languageId);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            if (lastError == Win32Errors.ERROR_INSUFFICIENT_BUFFER)
            {
               displayName = new StringBuilder((int) bufferSize + 1);

               goto Retry;
            }

            NativeError.ThrowException(lastError, _name);
         }


         return displayName.ToString();
      }


      /// <summary>この特権を表すために使用されるローカル一意識別子（LUID）を取得します（元のシステム上）。</summary>
      /// <returns>この特権を表すために使用されるローカル一意識別子（LUID）（元のシステム上）。</returns>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Luid")]
      [SecurityCritical]
      public long LookupLuid()
      {

         var success = NativeMethods.LookupPrivilegeValue(_systemName, _name, out var luid);

         var lastError = Marshal.GetLastWin32Error();
         if (!success)
         {
            NativeError.ThrowException(lastError, _name);
         }


         return Filesystem.NativeMethods.LuidToLong(luid);
      }


      /// <summary>特定の型のハッシュ関数として機能します。</summary>
      /// <returns>現在のオブジェクトのハッシュコード。</returns>
      public override int GetHashCode()
      {
         return !Utils.IsNullOrWhiteSpace(Name) ? Name.GetHashCode() : 0;
      }


      /// <summary>この特権のシステム名を返します。</summary>
      /// <remarks>これは<see cref="Privilege.Name"/>と同等です。</remarks>
      /// <returns>現在の<see cref="object"/>を表す<see cref="System.String"/>。</returns>
      public override string ToString()
      {
         return Name;
      }


      /// <summary>現在のオブジェクトが同じ型の別のオブジェクトと等しいかどうかを示します。</summary>
      /// <param name="other">このオブジェクトと比較するオブジェクト。</param>
      /// <returns>現在のオブジェクトが<paramref name="other"/>パラメーターと等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      public bool Equals(Privilege other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(Name, other.Name) &&
                Equals(_systemName, other._systemName);
      }


      /// <summary>指定された<see cref="object"/>が現在の<see cref="object"/>と等しいかどうかを判断します。</summary>
      /// <param name="obj">現在の<see cref="object"/>と比較する<see cref="object"/>。</param>
      /// <returns>指定された<see cref="object"/>が現在の<see cref="object"/>と等しい場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      /// <exception cref="NullReferenceException"/>
      public override bool Equals(object obj)
      {
         var other = obj as Privilege;

         return null != other && Equals(other);
      }


      /// <summary>==演算子を実装します。</summary>
      /// <param name="left">左辺の値。</param>
      /// <param name="right">右辺の値。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(Privilege left, Privilege right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>!=演算子を実装します。</summary>
      /// <param name="left">左辺の値。</param>
      /// <param name="right">右辺の値。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(Privilege left, Privilege right)
      {
         return !(left == right);
      }
      
      #endregion // Methods
   }
}
