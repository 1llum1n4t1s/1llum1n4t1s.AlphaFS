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
   /// <summary><see cref="BackupStreamInfo"/> 構造体はストリームヘッダーデータを格納します。</summary>
   /// <seealso cref="BackupFileStream"/>
   public sealed class BackupStreamInfo
   {
      #region Fields

      private readonly long _streamSize;
      private readonly string _streamName;
      private readonly StreamId _streamId;
      private readonly StreamAttribute _streamAttribute;

      #endregion // Fields


      #region Constructor

      /// <summary><see cref="BackupStreamInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="streamId">ストリームID。</param>
      /// <param name="name">名前。</param>
      internal BackupStreamInfo(NativeMethods.WIN32_STREAM_ID streamId, string name)
      {
         _streamName = name;
         _streamSize = (long)streamId.Size;
         _streamAttribute = (StreamAttribute)streamId.dwStreamAttribute;
         _streamId = (StreamId)streamId.dwStreamId;
      }

      #endregion // Constructor


      #region Public Properties

      /// <summary>サブストリーム内のデータサイズをバイト単位で取得します。</summary>
      /// <value>サブストリーム内のデータサイズ（バイト単位）。</value>
      public long Size
      {
         get { return _streamSize; }
      }


      /// <summary>代替データストリームの名前を指定する文字列を取得します。</summary>
      /// <value>代替データストリームの名前を指定する文字列。</value>
      public string Name
      {
         get { return _streamName; }
      }


      /// <summary>ストリーム内のデータの種類を取得します。</summary>
      /// <value>ストリーム内のデータの種類。</value>
      public StreamId StreamType
      {
         get { return _streamId; }
      }


      /// <summary>異なるオペレーティングシステム間での転送を容易にするためのデータ属性を取得します。</summary>
      /// <value>異なるオペレーティングシステム間での転送を容易にするためのデータ属性。</value>
      public StreamAttribute Attribute
      {
         get { return _streamAttribute; }
      }

      #endregion // Public Properties
   }
}
