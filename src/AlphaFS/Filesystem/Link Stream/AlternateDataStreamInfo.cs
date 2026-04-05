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
using System.Globalization;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>代替データストリームに関する情報を表します。</summary>
   /// <seealso cref="O:Alphaleonis.Win32.Filesystem.File.EnumerateAlternateDataStreams"/>
   [Serializable]
   public struct AlternateDataStreamInfo : IEquatable<AlternateDataStreamInfo>
   {
      #region Fields

      [NonSerialized] private readonly string _fullPath;
      [NonSerialized] private readonly string _streamName;
      
      #endregion // Fields


      #region Constructor

      internal AlternateDataStreamInfo(string fullPath, NativeMethods.WIN32_FIND_STREAM_DATA findData)
      {
         _fullPath = fullPath;

         Size = findData.StreamSize;

         _streamName = ParseStreamName(findData.cStreamName);
      }

      #endregion // Constructor


      #region Properties

      /// <summary>ストリームへのフルパスを取得します。</summary>
      /// <remarks>
      ///   これはロングパス形式のパスで、<see cref="PathFormat.FullPath"/> または <see cref="PathFormat.LongFullPath"/> を指定した場合に
      ///   <see cref="O:Alphaleonis.Win32.Filesystem.File.Open"/> に渡してストリームを開くことができます。
      /// </remarks>
      /// <value>ロングパス形式でのストリームへのフルパス。</value>
      public string FullPath
      {
         get { return string.Format(CultureInfo.InvariantCulture, "{0}{1}", _fullPath, !Utils.IsNullOrWhiteSpace(StreamName) ? Path.StreamSeparator + StreamName : string.Empty); }
      }
      

      /// <summary>ストリームのサイズを取得します。</summary>
      public long Size { get; private set; }


      /// <summary>代替データストリームの名前を取得します。</summary>
      /// <remarks>デフォルトストリーム (:$DATA) の場合は空文字列となり、その他のデータストリームの場合はストリーム名が含まれます。</remarks>
      /// <value>ストリームの名前。</value>
      public string StreamName
      {
         get { return _streamName; }
      }

      #endregion // Properties


      #region Methods

      /// <summary>このインスタンスのハッシュコードを返します。</summary>
      /// <returns>このインスタンスのハッシュコードである32ビット符号付き整数。</returns>
      public override int GetHashCode()
      {
         return Utils.CombineHashCodesOf(StreamName, FullPath);
      }
      

      /// <summary>指定した Object が現在の Object と等しいかどうかを判定します。</summary>
      /// <param name="other">比較対象の <see cref="AlternateDataStreamInfo"/> インスタンス。</param>
      /// <returns>指定した Object が現在の Object と等しい場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
      public bool Equals(AlternateDataStreamInfo other)
      {
         return GetType() == other.GetType() &&
                Equals(StreamName, other.StreamName) &&
                Equals(FullPath, other.FullPath) &&
                Equals(Size, other.Size);
      }


      /// <summary>このインスタンスと指定したオブジェクトが等しいかどうかを示します。</summary>
      /// <param name="obj">現在のインスタンスと比較するオブジェクト。</param>
      /// <returns>
      ///   <paramref name="obj"/> とこのインスタンスが同じ型で同じ値を表す場合は true、それ以外の場合は false。
      /// </returns>
      public override bool Equals(object obj)
      {
         return obj is AlternateDataStreamInfo && Equals((AlternateDataStreamInfo) obj);
      }


      // <summary>== 演算子を実装します。</summary>
      /// <param name="left">左辺の値。</param>
      /// <param name="right">右辺の値。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator ==(AlternateDataStreamInfo left, AlternateDataStreamInfo right)
      {
         return left.Equals(right);
      }


      /// <summary>!= 演算子を実装します。</summary>
      /// <param name="left">左辺の値。</param>
      /// <param name="right">右辺の値。</param>
      /// <returns>演算子の結果。</returns>
      public static bool operator !=(AlternateDataStreamInfo left, AlternateDataStreamInfo right)
      {
         return !(left == right);
      }

      #endregion // Methods


      #region Private Methods

      private static string ParseStreamName(string streamName)
      {
         if (null == streamName || streamName.Length < 2)
         {
            return string.Empty;
         }

         if (streamName[0] != Path.StreamSeparatorChar)
         {
            throw new ArgumentException(Resources.Invalid_Stream_Name, "streamName");
         }


         var sb = new StringBuilder(streamName.Length);

         for (int i = 1, l = streamName.Length; i < l; i++)
         {
            if (streamName[i] == Path.StreamSeparatorChar)
            {
               break;
            }

            sb.Append(streamName[i]);
         }


         return sb.ToString();
      }

      #endregion // Private Methods
   }
}
