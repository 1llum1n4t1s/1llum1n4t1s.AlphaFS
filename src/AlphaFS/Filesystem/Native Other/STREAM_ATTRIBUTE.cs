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
      /// <summary>異なるオペレーティングシステム間の転送を容易にする WIN32_STREAM_ID 構造体のデータ属性。このメンバーは次の値の 1 つ以上を指定できます。</summary>
      internal enum STREAM_ATTRIBUTE
      {
         /// <summary>このバックアップストリームには特別な属性がありません。</summary>
         NONE = 0,

         /// <summary>ストリームが読み取り時に変更されるデータを含む場合に設定される属性。バックアップアプリケーションがデータの検証が失敗することを認識できるようにします。</summary>
         STREAM_MODIFIED_WHEN_READ = 1,

         /// <summary>バックアップストリームにセキュリティ情報が含まれています。この属性は <see cref="STREAM_ID.BACKUP_SECURITY_DATA"/> 型のバックアップストリームにのみ適用されます。</summary>
         STREAM_CONTAINS_SECURITY = 2,

         /// <summary>予約済み。</summary>
         STREAM_CONTAINS_PROPERTIES = 4,

         /// <summary>バックアップストリームがスパースファイルストリームの一部です。この属性は <see cref="STREAM_ID.BACKUP_DATA"/>、<see cref="STREAM_ID.BACKUP_ALTERNATE_DATA"/>、および <see cref="STREAM_ID.BACKUP_SPARSE_BLOCK"/> 型のバックアップストリームにのみ適用されます。</summary>
         STREAM_SPARSE_ATTRIBUTE = 8
      }
   }
}
