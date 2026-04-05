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

namespace Alphaleonis.Win32.Security
{
   /// <summary>SECURITY_DESCRIPTOR_CONTROLデータ型は、セキュリティ記述子またはそのコンポーネントの意味を修飾するビットフラグのセットです。
   /// 各セキュリティ記述子には、SECURITY_DESCRIPTOR_CONTROLビットを格納するControlメンバーがあります。
   /// </summary>
   /// <remarks>
   /// <para>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</para>
   /// <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリのみ]</para>
   /// </remarks>
   [Flags]
   internal enum SECURITY_DESCRIPTOR_CONTROL
   {
      /// <summary>なし</summary>
      None = 0,

      /// <summary>SE_OWNER_DEFAULTED (0x0001) - デフォルトの所有者セキュリティ識別子（SID）を持つセキュリティ記述子を示します。このビットを使用して、デフォルトの所有者権限が設定されているすべてのオブジェクトを検索できます。</summary>
      SE_OWNER_DEFAULTED = 1,

      /// <summary>SE_GROUP_DEFAULTED (0x0002) - デフォルトのグループSIDを持つセキュリティ記述子を示します。このビットを使用して、デフォルトのグループ権限が設定されているすべてのオブジェクトを検索できます。</summary>
      SE_GROUP_DEFAULTED = 2,

      /// <summary>SE_DACL_PRESENT (0x0004) - 随意アクセス制御リスト（DACL）を持つセキュリティ記述子を示します。このフラグが設定されていない場合、またはこのフラグが設定されていてDACLがNULLの場合、セキュリティ記述子はすべてのユーザーにフルアクセスを許可します。</summary>
      SE_DACL_PRESENT = 4,

      /// <summary>SE_DACL_DEFAULTED (0x0008) - デフォルトのDACLを持つセキュリティ記述子を示します。例えば、オブジェクト作成者がDACLを指定しない場合、オブジェクトは作成者のアクセストークンからデフォルトのDACLを受け取ります。このフラグはACE継承に関してシステムがDACLを処理する方法に影響を与えることがあります。SE_DACL_PRESENTフラグが設定されていない場合、システムはこのフラグを無視します。</summary>
      SE_DACL_DEFAULTED = 8,

      /// <summary>SE_SACL_PRESENT (0x0010) - システムアクセス制御リスト（SACL）を持つセキュリティ記述子を示します。</summary>
      SE_SACL_PRESENT = 16,

      /// <summary>SE_SACL_DEFAULTED (0x0020) - デフォルトのSACLを持つセキュリティ記述子を示します。例えば、オブジェクト作成者がSACLを指定しない場合、オブジェクトは作成者のアクセストークンからデフォルトのSACLを受け取ります。このフラグはACE継承に関してシステムがSACLを処理する方法に影響を与えることがあります。SE_SACL_PRESENTフラグが設定されていない場合、システムはこのフラグを無視します。</summary>
      SE_SACL_DEFAULTED = 32,

      /// <summary>SE_DACL_AUTO_INHERIT_REQ (0x0100) - セキュリティ記述子で保護されたオブジェクトのプロバイダーが、DACLを既存の子オブジェクトに自動的に伝播することを要求します。プロバイダーが自動継承をサポートしている場合、DACLを既存の子オブジェクトに伝播し、オブジェクトとその子オブジェクトのセキュリティ記述子にSE_DACL_AUTO_INHERITEDビットを設定します。</summary>
      SE_DACL_AUTO_INHERIT_REQ = 256,

      /// <summary>SE_SACL_AUTO_INHERIT_REQ (0x0200) - セキュリティ記述子で保護されたオブジェクトのプロバイダーが、SACLを既存の子オブジェクトに自動的に伝播することを要求します。プロバイダーが自動継承をサポートしている場合、SACLを既存の子オブジェクトに伝播し、オブジェクトとその子オブジェクトのセキュリティ記述子にSE_SACL_AUTO_INHERITEDビットを設定します。</summary>
      SE_SACL_AUTO_INHERIT_REQ = 512,

      /// <summary>SE_DACL_AUTO_INHERITED (0x0400) - Windows 2000のみ。DACLが既存の子オブジェクトへの継承可能なACEの自動伝播をサポートするように設定されているセキュリティ記述子を示します。システムは、オブジェクトとその既存の子オブジェクトに対して自動継承アルゴリズムを実行するときにこのビットを設定します。</summary>
      SE_DACL_AUTO_INHERITED = 1024,

      /// <summary>SE_SACL_AUTO_INHERITED (0x0800) - Windows 2000: SACLが既存の子オブジェクトへの継承可能なACEの自動伝播をサポートするように設定されているセキュリティ記述子を示します。システムは、オブジェクトとその既存の子オブジェクトに対して自動継承アルゴリズムを実行するときにこのビットを設定します。</summary>
      SE_SACL_AUTO_INHERITED = 2048,

      /// <summary>SE_DACL_PROTECTED (0x1000) - Windows 2000: セキュリティ記述子のDACLが継承可能なACEによって変更されることを防止します。</summary>
      SE_DACL_PROTECTED = 4096,

      /// <summary>SE_SACL_PROTECTED (0x2000) - Windows 2000: セキュリティ記述子のSACLが継承可能なACEによって変更されることを防止します。</summary>
      SE_SACL_PROTECTED = 8192,

      /// <summary>SE_RM_CONTROL_VALID (0x4000) - リソースマネージャーコントロールが有効であることを示します。</summary>
      SE_RM_CONTROL_VALID = 16384,

      /// <summary>SE_SELF_RELATIVE (0x8000) - すべてのセキュリティ情報が連続したメモリブロックに含まれる自己相対形式のセキュリティ記述子を示します。このフラグが設定されていない場合、セキュリティ記述子は絶対形式です。詳細については、「絶対および自己相対セキュリティ記述子」を参照してください。</summary>
      SE_SELF_RELATIVE = 32768
   }
}
