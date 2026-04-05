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
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>コピーまたは移動操作の結果を格納するCopyMoveResultクラス。</summary>
   /// <remarks>通常、このクラスを手動でインスタンス化したり、値を設定する必要はありません。</remarks>
   [Serializable]
   public sealed class CopyMoveResult
   {
      #region Private Fields

      [NonSerialized] internal readonly Stopwatch Stopwatch;

      #endregion // Private Fields
      

      #region Constructors

      /// <summary>コピーまたは移動操作用のCopyMoveResultインスタンスを初期化します。</summary>
      /// <param name="source">ソースファイルまたはディレクトリのフルパス。</param>
      /// <param name="destination">宛先ファイルまたはディレクトリのフルパス。</param>
      private CopyMoveResult(string source, string destination)
      {
         Source = source;

         Destination = destination;

         IsCopy = true;

         Retries = 0;

         Stopwatch = new Stopwatch();
      }


      internal CopyMoveResult(CopyMoveArguments cma, bool isFolder) : this(cma.SourcePath, cma.DestinationPath)
      {
         IsEmulatedMove = cma.EmulateMove;

         IsCopy = cma.IsCopy;

         IsDirectory = isFolder;

         TimestampsCopied = cma.CopyTimestamps;
      }


      internal CopyMoveResult(CopyMoveArguments cma, bool isFolder, string source, string destination) : this(source, destination)
      {
         IsEmulatedMove = cma.EmulateMove;

         IsCopy = cma.IsCopy;

         IsDirectory = isFolder;

         TimestampsCopied = cma.CopyTimestamps;
      }

      #endregion // Constructors


      #region Properties

      /// <summary>コピーまたは移動操作の所要時間を示します。</summary>
      public TimeSpan Duration
      {
         get { return Stopwatch.Elapsed; }
      }
      

      /// <summary>宛先ファイルまたはディレクトリを示します。</summary>
      public string Destination { get; private set; }
      

      /// <summary>コピーまたは移動操作中に発生したエラーコード。</summary>
      /// <value>0（ゼロ）は成功を示します。</value>
      public int ErrorCode { get; internal set; }


      /// <summary>コピーまたは移動操作中に発生した<see cref="ErrorCode"/>からのエラーメッセージ。</summary>
      /// <value>エラーを説明するメッセージ。</value>
      [SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations")]
      public string ErrorMessage { get { return new Win32Exception(ErrorCode).Message; } }


      /// <summary><c>true</c>の場合、コピーまたは移動操作がキャンセルされたことを示します。</summary>
      /// <value>コピー/移動操作がキャンセルされた場合は<c>true</c>、それ以外の場合は<c>false</c>。</value>
      public bool IsCanceled { get; internal set; }


      /// <summary><c>true</c>の場合はコピー操作、それ以外は移動操作。</summary>
      /// <value>コピー操作の場合は<c>true</c>、移動操作の場合は<c>false</c>。</value>
      public bool IsCopy { get; private set; }


      /// <summary>このインスタンスがディレクトリを表すかどうかを示す値を取得します。</summary>
      /// <value>このインスタンスがディレクトリを表す場合は<c>true</c>、それ以外の場合は<c>false</c>。</value>
      public bool IsDirectory { get; private set; }


      /// <summary>移動操作がコピー＋削除のフォールバックを使用したことを示します。</summary>
      public bool IsEmulatedMove { get; private set; }


      /// <summary>このインスタンスがファイルを表すかどうかを示す値を取得します。</summary>
      /// <value>このインスタンスがファイルを表す場合は<c>true</c>、それ以外の場合は<c>false</c>。</value>
      public bool IsFile { get { return !IsDirectory; } }


      /// <summary><c>true</c>の場合は移動操作、それ以外はコピー操作。</summary>
      /// <value>移動操作の場合は<c>true</c>、コピー操作の場合は<c>false</c>。</value>
      public bool IsMove { get { return !IsCopy; } }


      /// <summary>リトライ試行の合計回数。</summary>
      public long Retries { get; internal set; }


      /// <summary>ソースファイルまたはディレクトリを示します。</summary>
      public string Source { get; private set; }


      /// <summary>ソースの日付とタイムスタンプが宛先ファイルシステムオブジェクトに適用されたことを示します。</summary>
      public bool TimestampsCopied { get; private set; }


      /// <summary>コピーされた合計バイト数。</summary>
      public long TotalBytes { get; internal set; }


      /// <summary>コピーされた合計バイト数（単位サイズとしてフォーマット済み）。</summary>
      public string TotalBytesUnitSize
      {
         get { return Utils.UnitSizeToText(TotalBytes); }
      }


      /// <summary>コピーされた合計ファイル数。</summary>
      public long TotalFiles { get; internal set; }


      /// <summary>コピーされた合計フォルダ数。</summary>
      public long TotalFolders { get; internal set; }

      #endregion // Properties
   }
}
