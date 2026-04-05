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
using System.Globalization;
namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>[AlphaFS] 既に存在するファイルまたはディレクトリを作成しようとしたときにスローされる例外。
   /// <c>ERROR_ALREADY_EXISTS</c>と<c>ERROR_FILE_EXISTS</c>の両方がこの例外を引き起こす可能性があります。
   /// </summary>
   [Serializable]
   public class AlreadyExistsException : System.IO.IOException
   {
      private static readonly int ErrorCode = Win32Errors.GetHrFromWin32Error(Win32Errors.ERROR_ALREADY_EXISTS);
      private static readonly string ErrorText = string.Format(CultureInfo.InvariantCulture, "({0}) {1}", Win32Errors.ERROR_ALREADY_EXISTS, new Win32Exception((int) Win32Errors.ERROR_ALREADY_EXISTS).Message.Trim().TrimEnd('.').Trim());


      /// <summary>[AlphaFS] <see cref="AlreadyExistsException"/>クラスの新しいインスタンスを初期化します。</summary>
      public AlreadyExistsException() : base(string.Format(CultureInfo.InvariantCulture, "{0}.", ErrorText), ErrorCode)
      {
      }


      /// <summary>[AlphaFS] <see cref="AlreadyExistsException"/>クラスの新しいインスタンスを初期化します。
      /// <c>ERROR_ALREADY_EXISTS</c>と<c>ERROR_FILE_EXISTS</c>の両方がこの例外を引き起こす可能性があります。
      /// </summary>
      /// <param name="message">カスタムエラーメッセージ。</param>
      public AlreadyExistsException(string message) : base(message, ErrorCode)
      {
      }


      /// <summary>[AlphaFS] <see cref="AlreadyExistsException"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">ファイルシステムオブジェクトへのパス。</param>
      /// <param name="isPath">このコンストラクタを使用する場合は常にtrueに設定します。</param>
      [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "isPath")]
      public AlreadyExistsException(string path, bool isPath) : base(string.Format(CultureInfo.InvariantCulture, "{0}: [{1}]", ErrorText, path), ErrorCode)
      {
      }


      /// <summary>[AlphaFS] <see cref="AlreadyExistsException"/>クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="path">ファイルシステムオブジェクトへのパス。</param>
      /// <param name="innerException">内部例外。</param>
      public AlreadyExistsException(string path, Exception innerException) : base(string.Format(CultureInfo.InvariantCulture, "{0}: [{1}]", ErrorText, path), innerException)
      {
      }
   }
}
