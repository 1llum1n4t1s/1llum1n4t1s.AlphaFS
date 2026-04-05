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
      /// <summary>バックアップストリームに含まれるデータの種類。</summary>
      internal enum STREAM_ID
      {
         /// <summary>エラーを示します。</summary>
         NONE = 0,

         /// <summary>標準データ。デフォルト（無名）データストリーム上の NTFS $DATA ストリームタイプに対応します。</summary>
         BACKUP_DATA = 1,

         /// <summary>拡張属性データ。NTFS $EA ストリームタイプに対応します。</summary>
         BACKUP_EA_DATA = 2,

         /// <summary>セキュリティ記述子データ。</summary>
         BACKUP_SECURITY_DATA = 3,

         /// <summary>代替データストリーム。名前付きデータストリーム上の NTFS $DATA ストリームタイプに対応します。</summary>
         BACKUP_ALTERNATE_DATA = 4,

         /// <summary>ハードリンク情報。NTFS $FILE_NAME ストリームタイプに対応します。</summary>
         BACKUP_LINK = 5,

         /// <summary>プロパティデータ。</summary>
         BACKUP_PROPERTY_DATA = 6,

         /// <summary>オブジェクト識別子。NTFS $OBJECT_ID ストリームタイプに対応します。</summary>
         BACKUP_OBJECT_ID = 7,

         /// <summary>リパースポイント。NTFS $REPARSE_POINT ストリームタイプに対応します。</summary>
         BACKUP_REPARSE_DATA = 8,

         /// <summary>スパースファイル。スパースファイル用の NTFS $DATA ストリームタイプに対応します。</summary>
         BACKUP_SPARSE_BLOCK = 9,

         /// <summary>トランザクション NTFS (TxF) データストリーム。</summary>
         /// <remarks>Windows Server 2003 および Windows XP: この値はサポートされていません。</remarks>
         BACKUP_TXFS_DATA = 10
      }
   }
}
