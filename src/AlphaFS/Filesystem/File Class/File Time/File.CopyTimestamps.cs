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
using System.IO;
using System.Security;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class File
   {
      #region Obsolete

      /// <summary>[AlphaFS] 指定されたファイルの日時スタンプを転送します。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [Obsolete("Use new method name: CopyTimestamp")]
      [SecurityCritical]
      public static void TransferTimestamps(string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         CopyTimestamps(sourcePath, destinationPath, pathFormat);
      }

      /// <summary>[AlphaFS] 指定されたファイルの日時スタンプを転送します。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>      
      [Obsolete("Use new method name: CopyTimestamp")]
      [SecurityCritical]
      public static void TransferTimestamps(string sourcePath, string destinationPath)
      {
         CopyTimestamps(sourcePath, destinationPath, PathFormat.RelativePath);
      }

      /// <summary>[AlphaFS] 指定されたファイルの日時スタンプを転送します。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>      
      [Obsolete("Use new method name: CopyTimestampsTransacted")]
      [SecurityCritical]
      public static void TransferTimestampsTransacted(KernelTransaction transaction, string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         CopyTimestampsTransacted(transaction, sourcePath, destinationPath, pathFormat);
      }

      /// <summary>[AlphaFS] 指定されたファイルの日時スタンプを転送します。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <param name="transaction">トランザクション。</param>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>      
      [Obsolete("Use new method name: CopyTimestampsTransacted")]
      [SecurityCritical]
      public static void TransferTimestampsTransacted(KernelTransaction transaction, string sourcePath, string destinationPath)
      {
         CopyTimestampsTransacted(transaction, sourcePath, destinationPath, PathFormat.RelativePath);
      }

      #endregion // Obsolete




      /// <summary>[AlphaFS] 指定された既存ファイルの日時とタイムスタンプをコピーします。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      [SecurityCritical]
      public static void CopyTimestamps(string sourcePath, string destinationPath)
      {
         CopyTimestampsCore(null, false, sourcePath, destinationPath, false, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された既存ファイルの日時とタイムスタンプをコピーします。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CopyTimestamps(string sourcePath, string destinationPath, PathFormat pathFormat)
      {
         CopyTimestampsCore(null, false, sourcePath, destinationPath, false, pathFormat);
      }


      /// <summary>[AlphaFS] 指定された既存ファイルの日時とタイムスタンプをコピーします。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      /// <param name="modifyReparsePoint"><c>true</c>の場合、日時情報はリパースポイント(シンボリックリンクまたはジャンクション)に適用され、リンク先のファイルには適用されません。<paramref name="destinationPath"/>がリパースポイントを参照していない場合は効果がありません。</param>
      [SecurityCritical]
      public static void CopyTimestamps(string sourcePath, string destinationPath, bool modifyReparsePoint)
      {
         CopyTimestampsCore(null, false, sourcePath, destinationPath, modifyReparsePoint, PathFormat.RelativePath);
      }


      /// <summary>[AlphaFS] 指定された既存ファイルの日時とタイムスタンプをコピーします。</summary>
      /// <remarks>このメソッドはソースファイルの最終アクセス日時を変更しません。</remarks>
      /// <exception cref="ArgumentException"/>
      /// <exception cref="FileNotFoundException"/>
      /// <exception cref="NotSupportedException"/>
      /// <param name="sourcePath">日時スタンプの取得元となるソースファイル。</param>
      /// <param name="destinationPath">日時スタンプを設定するコピー先ファイル。</param>
      /// <param name="modifyReparsePoint"><c>true</c>の場合、日時情報はリパースポイント(シンボリックリンクまたはジャンクション)に適用され、リンク先のファイルには適用されません。<paramref name="destinationPath"/>がリパースポイントを参照していない場合は効果がありません。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      [SecurityCritical]
      public static void CopyTimestamps(string sourcePath, string destinationPath, bool modifyReparsePoint, PathFormat pathFormat)
      {
         CopyTimestampsCore(null, false, sourcePath, destinationPath, modifyReparsePoint, pathFormat);
      }
   }
}
