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

namespace Alphaleonis.Win32.Security
{
   /// <summary>SE_OBJECT_TYPE列挙型は、セキュリティをサポートするWindowsオブジェクトの種類に対応する値を含みます。
   /// GetSecurityInfoやSetSecurityInfoなど、オブジェクトのセキュリティ情報を設定・取得する関数は、これらの値を使用してオブジェクトの種類を示します。
   /// </summary>
   /// <remarks>
   /// <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
   /// <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
   /// </remarks>
   internal enum SE_OBJECT_TYPE
   {
      /// <summary>不明なオブジェクトタイプ。</summary>
      SE_UNKNOWN_OBJECT_TYPE = 0,

      /// <summary>ファイルまたはディレクトリを示します。ファイルまたはディレクトリオブジェクトを識別する名前文字列は、以下の形式のいずれかです:
      ///   相対パス（例: FileName.dat または ..\FileName）
      ///   絶対パス（例: FileName.dat、C:\DirectoryName\FileName.dat、G:\RemoteDirectoryName\FileName.dat）
      ///   UNC名（例: \\ComputerName\ShareName\FileName.dat）
      /// </summary>
      SE_FILE_OBJECT,

      /// <summary>Windowsサービスを示します。サービスオブジェクトは、ServiceNameのようなローカルサービス、または\\ComputerName\ServiceNameのようなリモートサービスです。</summary>
      SE_SERVICE,

      /// <summary>プリンターを示します。プリンターオブジェクトは、PrinterNameのようなローカルプリンター、または\\ComputerName\PrinterNameのようなリモートプリンターです。</summary>
      SE_PRINTER,

      /// <summary>レジストリキーを示します。レジストリキーオブジェクトは、CLASSES_ROOT\SomePathのようなローカルレジストリ、または\\ComputerName\CLASSES_ROOT\SomePathのようなリモートレジストリに存在できます。
      /// レジストリキーの名前は、定義済みレジストリキーを識別するために"CLASSES_ROOT"、"CURRENT_USER"、"MACHINE"、"USERS"のリテラル文字列を使用する必要があります。
      /// </summary>
      SE_REGISTRY_KEY,

      /// <summary>ネットワーク共有を示します。共有オブジェクトは、ShareNameのようなローカル、または\\ComputerName\ShareNameのようなリモートです。</summary>
      SE_LMSHARE,

      /// <summary>ローカルカーネルオブジェクトを示します。GetSecurityInfoおよびSetSecurityInfo関数はすべてのカーネルオブジェクトをサポートします。
      /// GetNamedSecurityInfoおよびSetNamedSecurityInfo関数は、セマフォ、イベント、ミューテックス、待機タイマー、ファイルマッピングのカーネルオブジェクトのみで動作します。</summary>
      SE_KERNEL_OBJECT,

      /// <summary>ローカルコンピューター上のウィンドウステーションまたはデスクトップオブジェクトを示します。ウィンドウステーションまたはデスクトップの名前は一意ではないため、これらのオブジェクトにGetNamedSecurityInfoおよびSetNamedSecurityInfoは使用できません。</summary>
      SE_WINDOW_OBJECT,

      /// <summary>ディレクトリサービスオブジェクト、またはディレクトリサービスオブジェクトのプロパティセットまたはプロパティを示します。
      /// ディレクトリサービスオブジェクトの名前文字列はX.500形式である必要があります（例: CN=SomeObject,OU=ou2,OU=ou1,DC=DomainName,DC=CompanyName,DC=com,O=internet）。</summary>
      SE_DS_OBJECT,

      /// <summary>ディレクトリサービスオブジェクトとそのすべてのプロパティセットおよびプロパティを示します。</summary>
      SE_DS_OBJECT_ALL,

      /// <summary>プロバイダー定義オブジェクトを示します。</summary>
      SE_PROVIDER_DEFINED_OBJECT,

      /// <summary>WMIオブジェクトを示します。</summary>
      SE_WMIGUID_OBJECT,

      /// <summary>WOW64下のレジストリエントリのオブジェクトを示します。</summary>
      SE_REGISTRY_WOW64_32KEY
   }
}
