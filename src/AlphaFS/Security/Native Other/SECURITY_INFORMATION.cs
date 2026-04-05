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
   /// <summary>SECURITY_INFORMATIONデータ型は、設定または照会されるオブジェクト関連のセキュリティ情報を識別します。
   /// このセキュリティ情報には以下が含まれます:
   ///   オブジェクトの所有者;
   ///   オブジェクトのプライマリグループ;
   ///   オブジェクトの随意アクセス制御リスト（DACL）;
   ///   オブジェクトのシステムアクセス制御リスト（SACL）;
   /// </summary>
   /// <remarks>
   /// 符号なし32ビット整数がビットフラグによってSECURITY_DESCRIPTORの部分を指定します。
   /// 個々のビット値（ビットOR演算で組み合わせ可能）は以下の表の通りです。
   /// </remarks>
   [Flags]
   internal enum SECURITY_INFORMATION : uint
   {
      /// <summary>なし</summary>
      None = 0,

      /// <summary>OWNER_SECURITY_INFORMATION (0x00000001) - オブジェクトの所有者識別子が参照されています。</summary>
      OWNER_SECURITY_INFORMATION = 1,

      /// <summary>GROUP_SECURITY_INFORMATION (0x00000002) - オブジェクトのプライマリグループ識別子が参照されています。</summary>
      GROUP_SECURITY_INFORMATION = 2,

      /// <summary>DACL_SECURITY_INFORMATION (0x00000004) - オブジェクトのDACLが参照されています。</summary>
      DACL_SECURITY_INFORMATION = 4,

      /// <summary>SACL_SECURITY_INFORMATION (0x00000008) - オブジェクトのSACLが参照されています。</summary>
      SACL_SECURITY_INFORMATION = 8,

      /// <summary>LABEL_SECURITY_INFORMATION (0x00000010) - 必須整合性ラベルが参照されています。必須整合性ラベルはオブジェクトのSACL内のACEです。</summary>
      /// <remarks>Windows Server 2003およびWindows XP: このビットフラグは利用できません。</remarks>
      LABEL_SECURITY_INFORMATION = 16,

      /// <summary>ATTRIBUTE_SECURITY_INFORMATION (0x00000020) - 参照されるオブジェクトのリソースプロパティ。
      /// リソースプロパティはセキュリティ記述子のSACL内のSYSTEM_RESOURCE_ATTRIBUTE_ACE型に格納されます。
      /// </summary>
      /// <remarks>Windows Server 2008 R2、Windows 7、Windows Server 2008、Windows Vista、Windows Server 2003、Windows XP: このビットフラグは利用できません。</remarks>
      ATTRIBUTE_SECURITY_INFORMATION = 32,

      /// <summary>SCOPE_SECURITY_INFORMATION (0x00000040) - 参照されるオブジェクトに適用される集中アクセスポリシー（CAP）識別子。
      /// 各CAP識別子はセキュリティ記述子のSACL内のSYSTEM_SCOPED_POLICY_ID_ACE型に格納されます。
      /// </summary>
      /// <remarks>Windows Server 2008 R2、Windows 7、Windows Server 2008、Windows Vista、Windows Server 2003、Windows XP: このビットフラグは利用できません。</remarks>
      SCOPE_SECURITY_INFORMATION = 64,

      /// <summary>BACKUP_SECURITY_INFORMATION (0x00010000) - セキュリティ記述子のすべての部分。セキュリティ記述子全体を保持する必要があるバックアップおよび復元ソフトウェアに便利です。</summary>
      /// <remarks>Windows Server 2008 R2、Windows 7、Windows Server 2008、Windows Vista、Windows Server 2003、Windows XP: このビットフラグは利用できません。</remarks>
      BACKUP_SECURITY_INFORMATION = 65536,

      /// <summary>UNPROTECTED_SACL_SECURITY_INFORMATION (0x10000000) - SACLが親オブジェクトからACEを継承します。</summary>
      UNPROTECTED_SACL_SECURITY_INFORMATION = 268435456,

      /// <summary>UNPROTECTED_DACL_SECURITY_INFORMATION (0x20000000) - DACLが親オブジェクトからACEを継承します。</summary>
      UNPROTECTED_DACL_SECURITY_INFORMATION = 536870912,

      /// <summary>PROTECTED_SACL_SECURITY_INFORMATION (0x40000000) - SACLはACEを継承できません。</summary>
      PROTECTED_SACL_SECURITY_INFORMATION = 1073741824,

      /// <summary>PROTECTED_DACL_SECURITY_INFORMATION (0x80000000) - DACLはアクセス制御エントリ（ACE）を継承できません。</summary>
      PROTECTED_DACL_SECURITY_INFORMATION = 2147483648
   }
}
