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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>FINDEX_SEARCH_OPS 列挙型 - FindFirstFileEx 関数と共に使用して、実行するフィルタリングの種類を指定する値を定義します。</summary>
      /// <remarks>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      internal enum FINDEX_SEARCH_OPS
      {
         /// <summary>指定されたファイル名に一致するファイルを検索します。
         /// <para>この検索操作を使用する場合、FindFirstFileEx の lpSearchFilter パラメーターは NULL でなければなりません。</para>
         /// </summary>
         SearchNameMatch = 0,

         /// <summary>これはアドバイザリフラグです。ファイルシステムがディレクトリフィルタリングをサポートする場合、
         /// <para>関数は指定された名前に一致し、かつディレク��リであるファイルを検索します。</para>
         /// <para>ファイルシステムがディレクトリフィルタリングをサポートしない場合、このフラグは黙って無視されます。</para>
         /// <para>&#160;</para>
         /// <remarks>
         /// <para>この検索値を使用する場合、FindFirstFileEx 関数の lpSearchFilter パラメーターは NULL でなければなりません。</para>
         /// <para>ディレクトリフィルタリングが必要な場合、このフラグはすべてのファイルシステムで使用できますが、</para>
         /// <para>アドバイザリフラグであり、サポートするファイルシステムにのみ影響するため、</para>
         /// <para>アプリケーションは FindFirstFileEx 関数の lpFindFileData パラメーターに格納されたファイル属性データを</para>
         /// <para>調べて、関数がディレクトリへのハンドルを返したかどうかを判定する必要があります。</para>
         /// </remarks>
         /// </summary>
         SearchLimitToDirectories = 1,

         /// <summary>このフィルタリングタイプは利用できません。</summary>
         /// <remarks>詳細については、Device Interface Classes を参照してください。</remarks>
         SearchLimitToDevices = 2
      }
   }
}