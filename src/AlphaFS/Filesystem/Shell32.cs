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

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>Shell32を使用してファイルシステムオブジェクトへのアクセスを提供します。</summary>
   public static class Shell32
   {
      /// <summary>Shell32で使用されるIQueryAssociationsインターフェースメソッドの情報を提供します。</summary>
      [Flags]
      public enum AssociationAttributes
      {
         /// <summary>なし。</summary>
         None = 0,

         /// <summary>CLSID値をProgID値にマップしないよう指示します。</summary>
         InitNoRemapClsid = 1,

         /// <summary>指定されたファイルパラメータ（GetFileAssociation()関数の第3パラメータ）の値を実行可能ファイル名として識別します。</summary>
         /// <remarks>このフラグが設定されていない場合、ルートキーは実行可能ファイルのProgIDではなく、.exeキーに関連付けられたProgIDに設定されます。</remarks>
         InitByExeName = 2,

         /// <summary>IQueryAssociationメソッドがルートキーの下で要求された値を見つけられない場合、*サブキーから同等の値を取得しようとすることを指定します。</summary>
         InitDefaultToStar = 4,

         /// <summary>IQueryAssociationメソッドがルートキーの下で要求された値を見つけられない場合、Folderサブキーから同等の値を取得しようとすることを指定します。</summary>
         InitDefaultToFolder = 8,

         /// <summary>HKEY_CLASSES_ROOTのみを検索し、HKEY_CURRENT_USERを無視することを指定します。</summary>
         NoUserSettings = 16,

         /// <summary>返される文字列を切り詰めないことを指定します。代わりに、エラー値と完全な文字列に必要なサイズを返します。</summary>
         NoTruncate = 32,

         /// <summary>
         /// IQueryAssociationsメソッドにデータが正確であることを検証するよう指示します。
         /// この設定により、IQueryAssociationsメソッドは検証のためにユーザーのハードディスクからデータを読み取ることができます。
         /// 例えば、レジストリ内のフレンドリー名を.exeファイルに格納されているものと照合できます。
         /// </summary>
         /// <remarks>このフラグを設定すると、通常メソッドの効率が低下します。</remarks>
         Verify = 64,

         /// <summary>
         /// IQueryAssociationsメソッドにRundll.exeを無視してそのターゲットに関する情報を返すよう指示します。
         /// 通常、IQueryAssociationsメソッドはコマンド文字列内の最初の.exeまたは.dllに関する情報を返します。
         /// コマンドがRundll.exeを使用する場合、このフラグを設定するとメソッドはRundll.exeを無視してそのターゲットに関する情報を返します。
         /// </summary>
         RemapRunDll = 128,

         /// <summary>IQueryAssociationsメソッドに、関数のフレンドリー名が.exeファイルで見つかったものと一致しないなどのレジストリエラーを修正しないよう指示します。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "FixUps")]
         NoFixUps = 256,

         /// <summary>BaseClass値を無視することを指定します。</summary>
         IgnoreBaseClass = 512,

         /// <summary>「Unknown」ProgIDを無視し、代わりに失敗することを指定します。</summary>
         /// <remarks>Windows 7で導入されました。</remarks>
         InitIgnoreUnknown = 1024,

         /// <summary>指定されたProgIDがシステムのデフォルトを使用してマップされるべきであり、現在のユーザーのデフォルトではないことを指定します。</summary>
         /// <remarks>Windows 8で導入されました。</remarks>
         InitFixedProgId = 2048,

         /// <summary>値がプロトコルであり、現在のユーザーのデフォルトを使用してマップされるべきであることを指定します。</summary>
         /// <remarks>Windows 8で導入されました。</remarks>
         IsProtocol = 4096
      }


      //internal enum AssociationData
      //{
      //   MsiDescriptor = 1,
      //   NoActivateHandler = 2 ,
      //   QueryClassStore = 3,
      //   HasPerUserAssoc = 4,
      //   EditFlags = 5,
      //   Value = 6
      //}


      //internal enum AssociationKey
      //{
      //   ShellExecClass = 1,
      //   App = 2,
      //   Class = 3,
      //   BaseClass = 4
      //}


      /// <summary>ASSOCSTR列挙型 - AssocQueryString()関数が返す文字列の種類を定義するために使用されます。</summary>
      public enum AssociationString
      {
         /// <summary>なし。</summary>
         None = 0,

         /// <summary>シェル動詞に関連付けられたコマンド文字列。</summary>
         Command = 1,

         /// <summary>
         /// シェル動詞コマンド文字列からの実行可能ファイル。
         /// 例えば、この文字列はHKEY_CLASSES_ROOT\ApplicationName\shell\Open\commandなどのサブキーの（既定）値として見つかります。
         /// コマンドがRundll.exeを使用する場合、IQueryAssociations::GetStringのattributesパラメータに<see cref="AssociationAttributes.RemapRunDll"/>フラグを設定してターゲット実行可能ファイルを取得します。
         /// </summary>
         Executable = 2,

         /// <summary>ドキュメントタイプのフレンドリー名。</summary>
         FriendlyDocName = 3,

         /// <summary>実行可能ファイルのフレンドリー名。</summary>
         FriendlyAppName = 4,

         /// <summary>openサブキーに関連付けられた情報を無視します。</summary>
         NoOpen = 5,

         /// <summary>ShellNewサブキーの下を参照します。</summary>
         ShellNewValue = 6,

         /// <summary>DDEコマンドのテンプレート。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dde")]
         DdeCommand = 7,

         /// <summary>プロセスを作成するために使用するDDEコマンド。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dde")]
         DdeIfExec = 8,

         /// <summary>DDEブロードキャスト内のアプリケーション名。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dde")]
         DdeApplication = 9,

         /// <summary>DDEブロードキャスト内のトピック名。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dde")]
         DdeTopic = 10,

         /// <summary>
         /// InfoTipレジストリ値に対応します。
         /// アイテムのインフォチップ、またはファイル名上にカーソルを置いた時などにインフォチップを作成するためのIPropertyDescriptionList形式のプロパティリストを返します。
         /// プロパティリストはPSGetPropertyDescriptionListFromStringで解析できます。
         /// </summary>
         InfoTip = 11,

         /// <summary>
         /// QuickTipレジストリ値に対応します。<see cref="InfoTip"/>と同じですが、常にIPropertyDescriptionList形式のプロパティ名リストを返します。
         /// この値と<see cref="InfoTip"/>の違いは、オフラインや低速ネットワークなどプロパティ取得が遅いシナリオでも安全なプロパティを返す点です。
         /// <see cref="InfoTip"/>から返されるプロパティの一部は、低速プロパティ取得シナリオには適切でない場合があります。
         /// プロパティリストはPSGetPropertyDescriptionListFromStringで解析できます。
         /// </summary>
         QuickTip = 12,

         /// <summary>
         /// TileInfoレジストリ値に対応します。タイルビューのWindows Explorerウィンドウで特定のファイルタイプに対して表示するプロパティのリストを含みます。
         /// <see cref="InfoTip"/>と同じですが、<see cref="QuickTip"/>のようにIPropertyDescriptionList形式のプロパティ名リストも返します。
         /// プロパティリストはPSGetPropertyDescriptionListFromStringで解析できます。
         /// </summary>
         TileInfo = 13,

         /// <summary>
         /// imageやbmpなどのMIMEファイル関連付けの一般的な種類を記述し、
         /// アプリケーションが特定のファイルタイプについて一般的な仮定を行えるようにします。
         /// </summary>
         ContentType = 14,

         /// <summary>
         /// この関連付けにデフォルトで使用するアイコンリソースへのパスを返します。
         /// 正の数はDLLのリソーステーブルへのインデックスを示し、負の数はリソースIDを示します。
         /// リソースの構文例: "c:\myfolder\myfile.dll,-1"。
         /// </summary>
         DefaultIcon = 15,

         /// <summary>
         /// Shell拡張が関連付けられたオブジェクトに対して、IQueryAssociations::GetStringのpwszExtraパラメータとして
         /// 取得したいインターフェースのIIDの文字列表現を渡すことで、そのShell拡張オブジェクトのCLSIDを取得できます。
         /// 例えば、IExtractImageインターフェースを実装するハンドラーを取得する場合、
         /// IExtractImageのIIDである "{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}" を指定します。
         /// </summary>
         ShellExtension = 16,

         /// <summary>
         /// COMとIDropTargetインターフェースを介して呼び出される動詞に対して、このフラグを使用してIDropTargetオブジェクトのCLSIDを取得できます。
         /// このCLSIDはDropTargetサブキーに登録されています。
         /// 動詞はIQueryAssociations::GetStringの呼び出しで指定されたファイルパラメータで指定されます。
         /// </summary>
         DropTarget = 17,

         /// <summary>
         /// COMとIExecuteCommandインターフェースを介して呼び出される動詞に対して、このフラグを使用してIExecuteCommandオブジェクトのCLSIDを取得できます。
         /// このCLSIDは動詞のcommandサブキーにDelegateExecuteエントリとして登録されています。
         /// 動詞はIQueryAssociations::GetStringの呼び出しで指定されたファイルパラメータで指定されます。
         /// </summary>
         DelegateExecute = 18,

         /// <summary>（MSDNに説明はありません）</summary>
         /// <remarks>Windows 8で導入されました。</remarks>
         SupportedUriProtocols = 19,

         /// <summary>検証目的で使用される、<see cref="AssociationString"/>の定義済み最大値。</summary>
         Max = 20
      }


      /// <summary>Shell32 FileAttributes構造体。ファイルシステムオブジェクトの異なる種類を取得するために使用されます。</summary>
      [SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible")]
      [SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue")]
      [Flags]
      public enum FileAttributes
      {
         /// <summary>0x000000000 - ファイルシステムオブジェクトの大きいアイコンを取得します。</summary>
         /// <remarks><see cref="Icon"/>フラグも設定する必要があります。</remarks>
         LargeIcon = 0,

         /// <summary>0x000000001 - ファイルシステムオブジェクトの小さいアイコンを取得します。</summary>
         /// <remarks><see cref="Icon"/>フラグも設定する必要があります。</remarks>
         SmallIcon = 1,

         /// <summary>0x000000002 - ファイルシステムオブジェクトの開いた状態のアイコンを取得します。</summary>
         /// <remarks>コンテナオブジェクトはコンテナが開いていることを示すために開いたアイコンを表示します。</remarks>
         /// <remarks><see cref="Icon"/>および/または<see cref="SysIconIndex"/>フラグも設定する必要があります。</remarks>
         OpenIcon = 2,

         /// <summary>0x000000004 - ファイルシステムオブジェクトのShellサイズアイコンを取得します。</summary>
         /// <remarks>この属性が指定されていない場合、関数はシステムメトリック値に従ってアイコンのサイズを設定します。</remarks>
         ShellIconSize = 4,

         /// <summary>0x000000008 - PIDLでファイルシステムオブジェクトを取得します。</summary>
         /// <remarks>指定されたファイルがパス名ではなくITEMIDLIST構造体のアドレスを含むことを示します。</remarks>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Pidl")]
         Pidl = 8,

         /// <summary>0x000000010 - 指定されたファイルにアクセスすべきではないことを示します。代わりに、指定されたファイルが存在するかのように動作し、指定された属性を使用します。</summary>
         /// <remarks>このフラグは<see cref="Attributes"/>、<see cref="ExeType"/>、または<see cref="Pidl"/>属性と組み合わせることはできません。</remarks>
         UseFileAttributes = 16,

         /// <summary>0x000000020 - ファイルのアイコンに適切なオーバーレイを適用します。</summary>
         /// <remarks><see cref="Icon"/>フラグも設定する必要があります。</remarks>
         AddOverlays = 32,

         /// <summary>0x000000040 - オーバーレイアイコンのインデックスを返します。</summary>
         /// <remarks>オーバーレイインデックスの値は、psfiで指定された構造体のiIconメンバーの上位8ビットで返されます。</remarks>
         OverlayIndex = 64,

         /// <summary>0x000000100 - ファイルを表すアイコンのハンドルとシステムイメージリスト内のアイコンインデックスを取得します。ハンドルは構造体の<see cref="FileInfo.IconHandle"/>メンバーにコピーされ、インデックスは<see cref="FileInfo.IconIndex"/>メンバーにコピーされます。</summary>
         Icon = 256,

         /// <summary>0x000000200 - ファイルの表示名を取得します��名前は構造体の<see cref="FileInfo.DisplayName"/>メンバーにコピーされます。</summary>
         /// <remarks>返される表示名は、存在する場合は8.3形式ではなく長いファイル名を使用します。</remarks>
         DisplayName = 512,

         /// <summary>0x000000400 - ファイルの種類を説明する文字列を取得します。</summary>
         TypeName = 1024,

         /// <summary>0x000000800 - アイテム属性を取得します。属性は構造体の<see cref="FileInfo.Attributes"/>メンバーにコピーされます。</summary>
         /// <remarks>すべてのファイルにアクセスするため、パフォーマンスが低下します。</remarks>
         Attributes = 2048,

         /// <summary>0x000001000 - pszPathで指定されたファイルを表すアイコンを含むファイルの名前を取得します。アイコンを含むファイル名は構造体の<see cref="FileInfo.DisplayName"/>メンバーにコピーされ、アイコンのインデックスは<see cref="FileInfo.IconIndex"/>メンバーにコピーされます。</summary>
         IconLocation = 4096,

         /// <summary>0x000002000 - pszPathが実行可能ファイルを識別する場合、実行可能ファイルの種類を取得します。</summary>
         /// <remarks>このフラグは他の属性と一緒に指定できません。</remarks>
         ExeType = 8192,

         /// <summary>0x000004000 - システムイメージリストアイコンのインデックスを取得します。</summary>
         SysIconIndex = 16384,

         /// <summary>0x000008000 - ファイルのアイコンにリンクオーバーレイを追加します。</summary>
         /// <remarks><see cref="Icon"/>フラグも設定する必要があります。</remarks>
         LinkOverlay = 32768,

         /// <summary>0x000010000 - ファイルのアイコンをシステムのハイライト色とブレンドします。</summary>
         Selected = 65536,

         /// <summary>0x000020000 - <see cref="FileInfo.Attributes"/>が要求される特定の属性を含むことを示すために<see cref="Attributes"/>を変更します。</summary>
         /// <remarks>このフラグは<see cref="Icon"/>属性と一緒に指定できません。すべてのファイルにアクセスするため、パフォーマンスが低下します。</remarks>
         AttributesSpecified = 131072
      }


      /// <summary>SHFILEINFO構造体。ファイルシステムオブジェクトに関する情報を格納します。</summary>
      [SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes")]
      [SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible")]
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      public struct FileInfo
      {
         /// <summary>ファイルを表すアイコンへのハンドル。</summary>
         /// <remarks>呼び出し元は不要になった時点でDestroyIcon()で��のハンドルを破棄する責任があります。</remarks>
         [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
         public readonly IntPtr IconHandle;

         /// <summary>システムイメージリスト内のアイコン画像のインデックス。</summary>
         [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
         public int IconIndex;

         /// <summary>ファイルオブジェクトの属性を示す値の配列。</summary>
         [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
         [MarshalAs(UnmanagedType.U4)]
         public readonly GetAttributesOf Attributes;

         /// <summary>Windows Shellに表示されるファイル名、またはファイルを表すアイコンを含��ファイルのパスとファイル名。</summary>
         [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = NativeMethods.MaxPath)]
         public string DisplayName;

         /// <summary>ファイルの種類。</summary>
         [SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")]
         [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
         public string TypeName;
      }


      /// <summary>SFGAO - ファイルシステムオブジェクトから取得できる属性。</summary>
      [SuppressMessage("Microsoft.Usage", "CA2217:DoNotMarkEnumsWithFlags"), SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible")]
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Sh")]
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Sh")]
      [SuppressMessage("Microsoft.Naming", "CA1714:FlagsEnumsShouldHavePluralNames")]
      [Flags]
      public enum GetAttributesOf
      {
         /// <summary>0x00000000 - なし。</summary>
         None = 0,

         /// <summary>0x00000001 - 指定されたアイテムをコピーできます。</summary>
         CanCopy = 1,

         /// <summary>0x00000002 - 指定されたアイテムを移動できます。</summary>
         CanMove = 2,

         /// <summary>0x00000004 - 指定されたアイテムのショートカットを作成できます。</summary>
         CanLink = 4,

         /// <summary>0x00000008 - 指定されたアイテムをIShellFolder::BindToObjectを通じてIStorageオブジェクトにバインドできます。名前空間操作機能の詳細はIStorageを参照してください。</summary>
         Storage = 8,

         /// <summary>0x00000010 - 指定されたアイテムの名前を変更できます。この値は本質的に提案であり、すべての名前空間クライアントがアイテムの名前変更を許可するわけではありません。ただし、許可するものはこの属性を設定する必要があります。</summary>
         CanRename = 16,

         /// <summary>0x00000020 - 指定されたアイテムを削除できます。</summary>
         CanDelete = 32,

         /// <summary>0x00000040 - 指定されたアイテムにはプロパティシートがあります。</summary>
         HasPropSheet = 64,

         /// <summary>0x00000100 - 指定されたアイテムはドロップターゲットです。</summary>
         DropTarget = 256,

         /// <summary>0x00001000 - 指定されたアイテムはシステムアイテムです。</summary>
         ///  <remarks>Windows 7以降。</remarks>
         System = 4096,

         /// <summary>0x00002000 - 指定されたアイテムは暗号化されており、特別な表示が必要な場合があります。</summary>
         Encrypted = 8192,

         /// <summary>0x00004000 - アイテムへのアクセス（IStreamまたは他のストレージインターフェースを介して）は低速な操作になることが予想されます。</summary>
         IsSlow = 16384,

         /// <summary>0x00008000 - 指定されたアイテムは淡色表示され、ユーザーが使用できない状態で表示されます。</summary>
         Ghosted = 32768,

         /// <summary>0x00010000 - 指定されたアイテムはショートカットです。</summary>
         Link = 65536,

         /// <summary>0x00020000 - 指定されたオブジェクトは共有されています。</summary>
         Share = 131072,

         /// <summary>0x00040000 - 指定されたアイテムは読み取り専用です。フォルダの場合、そのフォルダに新しいアイテムを作成できないことを意味します。</summary>
         ReadOnly = 262144,

         /// <summary>0x00080000 - アイテムは隠しファイルであり、フォルダ設定で「隠しファイルとフォルダの表示」オプションが有効でない限り表示されません。</summary>
         Hidden = 524288,

         /// <summary>0x00100000 - アイテムは列挙されないアイテムであり、非表示にする必要があります。IShellFolder::EnumObjectsメソッドで作成されたものなどの列挙子を介して返されません。</summary>
         NonEnumerated = 1048576,

         /// <summary>0x00200000 - 特定のアプリケーションで定義された新しいコンテンツがアイテムに含まれています。</summary>
         NewContent = 2097152,

         /// <summary>0x00400000 - アイテムに関連付けられたストリームがあることを示します。</summary>
         Stream = 4194304,

         /// <summary>0x00800000 - このアイテムの子はIStreamまたはIStorageを通じてアクセスできます。</summary>
         StorageAncestor = 8388608,

         /// <summary>0x01000000 - 入力として指定された場合、フォルダまたはShellアイテム配列に含まれるアイテムが存在することを検証するようフォルダに指示します。</summary>
         Validate = 16777216,

         /// <summary>0x02000000 - 指定されたアイテムはリムーバブルメディア上にあるか、リムーバブルデバイスそのものです。</summary>
         Removable = 33554432,

         /// <summary>0x04000000 - 指定されたアイテムは圧縮されています。</summary>
         Compressed = 67108864,

         /// <summary>0x08000000 - 指定されたアイテムはWebブラウザまたはWindows Explorerフレーム内でホストできます。</summary>
         Browsable = 134217728,

         /// <summary>0x10000000 - 指定されたフォルダはファイルシステムフォルダであるか、ファイルシステムフォルダである少なくとも1つの子孫（子、孫、またはそれ以降）を含みます。</summary>
         FileSysAncestor = 268435456,

         /// <summary>0x20000000 - 指定されたアイテムはフォルダです。</summary>
         Folder = 536870912,

         /// <summary>0x40000000 - 指定されたフォルダまたはファイルはファイルシステムの一部です（つまり、ファイル、ディレクトリ、またはルートディレクトリです）。</summary>
         FileSystem = 1073741824,

         /// <summary>0x80000000 - 指定されたフォルダにはサブフォルダがあります。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1702:CompoundWordsShouldBeCasedCorrectly", MessageId = "SubFolder")]
         HasSubFolder = unchecked((int)0x80000000)
      }


      /// <summary>UrlIs()メソッドでURLの種類を定義するために使用されます。</summary>
      public enum UrlType
      {
         /// <summary>URLは有効ですか？</summary>
         IsUrl = 0,

         /// <summary>URLは不透明ですか？</summary>
         IsOpaque = 1,

         /// <summary>URLはナビゲーション履歴で通常追跡されないURLですか？</summary>
         IsNoHistory = 2,

         /// <summary>URLはファイルURLですか？</summary>
         IsFileUrl = 3,

         /// <summary>URLの有効なスキームを特定しようとします。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Appliable")]
         IsAppliable = 4,

         /// <summary>URL文字列はディレクトリで終わっていますか？</summary>
         IsDirectory = 5,

         /// <summary>URLに追加されたクエリ文字列がありますか？</summary>
         IsHasQuery = 6
      }


      #region Methods

      /// <summary>アイコンを破棄し、アイコンが占有していたメモリを解放します。</summary>
      /// <param name="iconHandle">アイコンへの<see cref="IntPtr"/>ハンドル。</param>
      public static void DestroyIcon(IntPtr iconHandle)
      {
         if (IntPtr.Zero != iconHandle)
         {
            NativeMethods.DestroyIcon(iconHandle);
         }
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたファイルまたはプロトコルを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連の文字列。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileAssociation(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.Executable);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたコンテンツタイプを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連のコンテンツタイプ。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileContentType(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.ContentType);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたデフォルトアイコンを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連のデフォルトアイコン。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileDefaultIcon(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.DefaultIcon);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたフレンドリーなアプリケーション名を取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連のフレンドリーなアプリケーション名。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileFriendlyAppName(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.InitByExeName, AssociationString.FriendlyAppName);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたフレンドリーなドキュメント名を取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連のフレンドリーなドキュメント名。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileFriendlyDocName(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.FriendlyDocName);
      }


      /// <summary>ファイルを表すShellアイコンへの<see cref="IntPtr"/>ハンドルを取得します。</summary>
      /// <remarks>呼び出し元は不要になった時点でDestroyIcon()でこのハンドルを破棄する責任があります。</remarks>
      /// <param name="filePath">
      ///   最大パス長を超えないファイルシステムオブジェクトへのパス。絶対パスと相対パスの両方が有効です。
      /// </param>
      /// <param name="iconAttributes">
      ///   アイコンサイズ <see cref="Shell32.FileAttributes.SmallIcon"/> または <see cref="Shell32.FileAttributes.LargeIcon"/>。
      ///   <see cref="Shell32.FileAttributes.AddOverlays"/>などと組み合わせることもできます。
      /// </param>
      /// <returns>ファイルを表すShellアイコンへの<see cref="IntPtr"/>ハンドル。失敗時はIntPtr.Zero。</returns>
      [SecurityCritical]
      public static IntPtr GetFileIcon(string filePath, FileAttributes iconAttributes)
      {
         if (Utils.IsNullOrWhiteSpace(filePath))
         {
            return IntPtr.Zero;
         }

         var fileInfo = GetFileInfoCore(filePath, System.IO.FileAttributes.Normal, FileAttributes.Icon | iconAttributes, true, true);
         return fileInfo.IconHandle == IntPtr.Zero ? IntPtr.Zero : fileInfo.IconHandle;
      }


      /// <summary>ファイル、フォルダ、ディレクトリ、ドライブルートなどのファイルシステム内のオブジェクトに関する情報を取得します。</summary>
      /// <returns><see cref="FileInfo"/>構造体インスタンス。</returns>
      /// <remarks>
      /// <para>この関数はバックグラウンドスレッドから呼び出す必要があります。</para>
      /// <para>そうしないと、UIが応答を停止する可能性があります。</para>
      /// <para>Unicodeパスがサポートされています。</para>
      /// </remarks>
      /// <param name="filePath">最大パス長を超えないファイルシステムオブジェクトへのパス。絶対パスと相対パスの両方が有効です。</param>
      /// <param name="attributes"><see cref="System.IO.FileAttributes"/>属性。</param>
      /// <param name="fileAttributes">1つ以上の<see cref="FileAttributes"/>属性。</param>
      /// <param name="continueOnException">
      /// <para><c>true</c>の場合、ACL保護されたディレクトリやアクセスできないリパースポイントなどの障害の結果としてスローされる例外を抑制します。</para>
      /// </param>
      [SecurityCritical]
      public static FileInfo GetFileInfo(string filePath, System.IO.FileAttributes attributes, FileAttributes fileAttributes, bool continueOnException)
      {
         return GetFileInfoCore(filePath, attributes, fileAttributes, true, continueOnException);
      }


      /// <summary>指定されたファイルに関する情報を含む<see cref="Shell32Info"/>のインスタンスを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns><see cref="Shell32Info"/>クラスのインスタンス。</returns>
      [SecurityCritical]
      public static Shell32Info GetShell32Info(string path)
      {
         return new Shell32Info(path);
      }

      /// <summary>指定されたファイルに関する情報を含む<see cref="Shell32Info"/>のインスタンスを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      /// <returns><see cref="Shell32Info"/>クラスのインスタンス。</returns>
      [SecurityCritical]
      public static Shell32Info GetShell32Info(string path, PathFormat pathFormat)
      {
         return new Shell32Info(path, pathFormat);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられた「プログラムから開く」コマンドを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連の「プログラムから開く」コマンド。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileOpenWithAppName(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.FriendlyAppName);
      }


      /// <summary>レジストリから<paramref name="path"/>に関連付けられたShellコマンドを取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連のShellコマンド。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      [SecurityCritical]
      public static string GetFileVerbCommand(string path)
      {
         return GetFileAssociationCore(path, AssociationAttributes.Verify, AssociationString.Command);
      }


      /// <summary>ファイルURLをMicrosoft MS-DOSパスに変換します。</summary>
      /// <param name="urlPath">ファイルURL。</param>
      /// <returns>
      /// <para>Microsoft MS-DOSパス。パスが作成できない場合は<c>string.Empty</c>が返されます。</para>
      /// <para><paramref name="urlPath"/>が<c>null</c>の場合は<c>null</c>が返されます。</para>
      /// </returns>
      [SecurityCritical]
      internal static string PathCreateFromUrl(string urlPath)
      {
         if (urlPath == null)
         {
            return null;
         }

         var buffer = new StringBuilder(NativeMethods.MaxPathUnicode);
         var bufferSize = (uint)buffer.Capacity;

         var lastError = NativeMethods.PathCreateFromUrl(urlPath, buffer, ref bufferSize, 0);

         // 例外をスローせず、string.Emptyを返します。
         return lastError == Win32Errors.S_OK ? buffer.ToString() : string.Empty;
      }


      /// <summary>ファイルURLからパスを作成します。</summary>
      /// <returns>
      /// <para>ファイルパス。パスが作成できない場合は<c>string.Empty</c>が返されます。</para>
      /// <para><paramref name="urlPath"/>が<c>null</c>の場合は<c>null</c>が返されます。</para>
      /// </returns>
      /// <exception cref="PlatformNotSupportedException">オペレーティングシステムがWindows Vistaより古い場合。</exception>
      /// <param name="urlPath">URL。</param>
      [SecurityCritical]
      internal static string PathCreateFromUrlAlloc(string urlPath)
      {
         if (!NativeMethods.IsAtLeastWindowsVista)
         {
            throw new PlatformNotSupportedException(new Win32Exception((int)Win32Errors.ERROR_OLD_WIN_VERSION).Message);
         }


         if (urlPath == null)
         {
            return null;
         }

         var lastError = NativeMethods.PathCreateFromUrlAlloc(urlPath, out var buffer, 0);

         // 例外をスローせず、string.Emptyを返します。
         return lastError == Win32Errors.S_OK ? buffer.ToString() : string.Empty;
      }


      /// <summary>ファイルやフォルダなどのファイルシステムオブジェクトへのパスが有効かどうかを判断します。</summary>
      /// <param name="path">検証するオブジェクトへの最大パス長のフルパス。</param>
      /// <returns>ファイルが存在する場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      [SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "lastError")]
      [SecurityCritical]
      public static bool PathFileExists(string path)
      {
         // PathFileExists()
         // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。

         return !Utils.IsNullOrWhiteSpace(path) && NativeMethods.PathFileExists(Path.GetFullPathCore(null, false, path, GetFullPathOptions.AsLongPath | GetFullPathOptions.FullCheck | GetFullPathOptions.ContinueOnNonExist));
      }


      /// <summary>URLが指定された種類であるかどうかをテストします。</summary>
      /// <param name="url">URL。</param>
      /// <param name="urlType"></param>
      /// <returns>
      /// URLタイプの1つを除くすべてについて、URLが指定された種類の場合はUrlIsは<c>true</c>を返し、それ以外の場合は<c>false</c>を返します。
      /// UrlIsが<see cref="UrlType.IsAppliable"/>に設定されている場合、UrlIsはURLスキームの特定を試みます。
      /// 関数がスキームを特定できた場合は<c>true</c>を返し、そうでなければ<c>false</c>を返します。
      /// </returns>
      [SecurityCritical]
      internal static bool UrlIs(string url, UrlType urlType)
      {
         return NativeMethods.UrlIs(url, urlType);
      }


      /// <summary>Microsoft MS-DOSパスを正規化されたURLに変換します。</summary>
      /// <param name="path">最大長<see cref="NativeMethods.MaxPath"/>の完全なMS-DOSパス。</param>
      /// <returns>
      /// <para>URL。URLが作成できない場合は<c>string.Empty</c>が返されます。</para>
      /// <para><paramref name="path"/>が<c>null</c>の場合は<c>null</c>が返されます。</para>
      /// </returns>
      [SecurityCritical]
      internal static string UrlCreateFromPath(string path)
      {
         if (path == null)
         {
            return null;
         }

         // UrlCreateFromPathは拡張パスをサポートしていません。
         var pathRp = Path.GetRegularPathCore(path, GetFullPathOptions.CheckInvalidPathChars, false);

         var buffer = new StringBuilder(NativeMethods.MaxPathUnicode);
         var bufferSize = (uint)buffer.Capacity;

         var lastError = NativeMethods.UrlCreateFromPath(pathRp, buffer, ref bufferSize, 0);

         // 例外をスローせず、nullを返します。
         var url = buffer.ToString();
         if (Utils.IsNullOrWhiteSpace(url))
         {
            url = string.Empty;
         }

         return lastError == Win32Errors.S_OK ? url : string.Empty;
      }


      /// <summary>URLがファイルURLであるかどうかをテストします。</summary>
      /// <param name="url">URL。</param>
      /// <returns>URLがファイルURLの場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      [SecurityCritical]
      internal static bool UrlIsFileUrl(string url)
      {
         return NativeMethods.UrlIs(url, UrlType.IsFileUrl);
      }


      /// <summary>URLがブラウザのナビゲーション履歴に通常含まれないURLであるかどうかを返します。</summary>
      /// <param name="url">URL。</param>
      /// <returns>URLがナビゲーション履歴に含まれないURLの場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      [SecurityCritical]
      internal static bool UrlIsNoHistory(string url)
      {
         return NativeMethods.UrlIs(url, UrlType.IsNoHistory);
      }


      /// <summary>URLが不透明であるかどうかを返します。</summary>
      /// <param name="url">URL。</param>
      /// <returns>URLが不透明な場合は<c>true</c>、それ以外の場合は<c>false</c>。</returns>
      [SecurityCritical]
      internal static bool UrlIsOpaque(string url)
      {
         return NativeMethods.UrlIs(url, UrlType.IsOpaque);
      }


      #region Internal Methods

      /// <summary>レジストリからファイルまたはプロトコルの関連付けに関連する文字列を検索して取得します。</summary>
      /// <param name="path">ファイルへのパス。</param>
      /// <param name="attributes">1つ以上の<see cref="AssociationAttributes"/>属性。「InitXXX」属性は1つだけ使用できます。</param>
      /// <param name="associationType"><see cref="AssociationString"/>属性。</param>
      /// <returns>レジストリから関連するファイルまたはプロトコル関連の文字列。関連付けが見つからない場合は<c>string.Empty</c>。</returns>
      /// <exception cref="ArgumentNullException"/>
      [SecurityCritical]
      private static string GetFileAssociationCore(string path, AssociationAttributes attributes, AssociationString associationType)
      {
         if (Utils.IsNullOrWhiteSpace(path))
         {
            throw new ArgumentNullException("path");
         }

         attributes = attributes | AssociationAttributes.NoTruncate | AssociationAttributes.RemapRunDll;

         uint bufferSize = NativeMethods.MaxPath;
         StringBuilder buffer;
         uint retVal;

         do
         {
            buffer = new StringBuilder((int)bufferSize);

            // AssocQueryString()
            // 2014-02-05: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。
            // 2015-07-17: この関数は長いパスをサポートしていません。

            retVal = NativeMethods.AssocQueryString(attributes, associationType, path, null, buffer, out bufferSize);

            // エラー時に例外はスローされず、空の文字列を返します。

            //switch (retVal)
            //{
            //   // 0x80070483: No application is associated with the specified file for this operation.
            //   case 2147943555:
            //   case Win32Errors.E_POINTER:
            //   case Win32Errors.S_OK:
            //      break;

            //   default:
            //      NativeError.ThrowException(retVal);
            //      break;
            //}

         } while (retVal == Win32Errors.E_POINTER);

         return buffer.ToString();
      }


      /// <summary>ファイル、フォルダ、ディレクトリ、ドライブルートなどのファイルシステム内のオブジェクトに関する情報を取得します。</summary>
      /// <returns><see cref="FileInfo"/>構造体インスタンス。</returns>
      /// <remarks>
      /// <para>この関数はバックグラウンドスレッドから呼び出す必要があります。</para>
      /// <para>そうしないと、UIが応答を停止する可能性があります。</para>
      /// <para>Unicodeパスはサポートされていません。</para>
      /// </remarks>
      /// <param name="path">最大パス長を超えないファイルシステムオブジェクトへのパス。絶対パスと相対パスの両方が有効です。</param>
      /// <param name="attributes"><see cref="System.IO.FileAttributes"/>属性。</param>
      /// <param name="fileAttributes"><see cref="FileAttributes"/>属性。</param>
      /// <param name="checkInvalidPathChars">パスに有効なパス文字のみが含まれているかどうかを確認します。</param>
      /// <param name="continueOnException">
      /// <para><c>true</c>の場合、ACL保護されたディレクトリやアクセスできないリパースポイントなどの障害の結果としてスローされる例外を抑制します。</para>
      /// </param>
      [SecurityCritical]
      internal static FileInfo GetFileInfoCore(string path, System.IO.FileAttributes attributes, FileAttributes fileAttributes, bool checkInvalidPathChars, bool continueOnException)
      {
         // クラッシュの可能性を防止します。
         var fileInfo = new FileInfo
         {
            DisplayName = string.Empty,
            TypeName = string.Empty,
            IconIndex = 0
         };

         if (!Utils.IsNullOrWhiteSpace(path))
         {
            // ShGetFileInfo()
            // 2013-01-13: MSDNはLongPathの使用を確認していませんが、この関数のUnicodeバージョンが存在します。
            // 2015-07-17: この関数は長いパスをサポートしていません。

            var shGetFileInfo = NativeMethods.ShGetFileInfo(Path.GetRegularPathCore(path, checkInvalidPathChars ? GetFullPathOptions.CheckInvalidPathChars : 0, false), attributes, out fileInfo, (uint)Marshal.SizeOf(fileInfo), fileAttributes);

            if (shGetFileInfo == IntPtr.Zero && !continueOnException)
            {
               NativeError.ThrowException(Marshal.GetLastWin32Error(), path);
            }
         }

         return fileInfo;
      }

      #endregion // Internal Methods

      #endregion // Methods
   }
}
