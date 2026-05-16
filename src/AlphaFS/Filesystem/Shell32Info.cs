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
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ファイルに関するShell32情報を格納します。</summary>
   [Serializable]
   [SecurityCritical]
   public sealed class Shell32Info : IDisposable
   {
      #region Constructors

      /// <summary>Shell32Infoインスタンスを初期化します。</summary>
      /// <remarks>Shell32は<c>MAX_PATH</c>の長さに制限されています。</remarks>
      /// <remarks>このコンストラクタはファイルの存在を確認しません。後続の操作でファイルにアクセスするための文字列のプレースホルダーです。</remarks>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      public Shell32Info(string fileName) : this(fileName, PathFormat.RelativePath)
      {
      }

      /// <summary>Shell32Infoインスタンスを初期化します。</summary>
      /// <remarks>Shell32は<c>MAX_PATH</c>の長さに制限されています。</remarks>
      /// <remarks>このコンストラクタはファイルの存在を確認しません。後続の操作でファイルにアクセスするための文字列のプレースホルダーです。</remarks>
      /// <param name="fileName">新しいファイルの完全修飾名、または相対ファイル名。パスの末尾にディレクトリ区切り文字を付けないでください。</param>
      /// <param name="pathFormat">パスパラメータの形式を示します。</param>
      public Shell32Info(string fileName, PathFormat pathFormat)
      {
         if (Utils.IsNullOrWhiteSpace(fileName))
         {
            throw new ArgumentNullException("fileName");
         }

         // Shell32は MAX_PATH の長さに制限されています。
         // 通常形式のフルパスを取得します。

         FullPath = Path.GetExtendedLengthPathCore(null, fileName, pathFormat, GetFullPathOptions.RemoveTrailingDirectorySeparator | GetFullPathOptions.FullCheck);

         Initialize();
      }

      #endregion // Constructors

      
      #region Methods

      /// <summary>ファイルを表すShellアイコンへの<see cref="IntPtr"/>ハンドルを取得します。</summary>
      /// <param name="iconAttributes">アイコンサイズ <see cref="Shell32.FileAttributes.SmallIcon"/> または <see cref="Shell32.FileAttributes.LargeIcon"/>。<see cref="Shell32.FileAttributes.AddOverlays"/>などと組み合わせることもできます。</param>
      /// <returns>ファイルを表すShellアイコンへの<see cref="IntPtr"/>ハンドル。</returns>
      /// <remarks>呼び出し元は不要になった時点でDestroyIcon()でこのハンドルを破棄する責任があります。</remarks>
      [SecurityCritical]
      public IntPtr GetIcon(Shell32.FileAttributes iconAttributes)
      {
         return Shell32.GetFileIcon(FullPath, iconAttributes);
      }


      /// <summary>レジストリからShellコマンドの関連付けを取得します。</summary>
      /// <param name="shellVerb">シェル動詞。</param>
      /// <returns>
      ///   レジストリから関連するファイルまたはプロトコル関連のShellコマンドを返します。関連付けが見つからない場合は<c>string.Empty</c>を返します。
      /// </returns>
      [SecurityCritical]
      public string GetVerbCommand(string shellVerb)
      {
         return GetString(_iQaNone, Shell32.AssociationString.Command, shellVerb);
      }


      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      private static string GetString(NativeMethods.QueryAssociationsWrapper iQa, Shell32.AssociationString assocString, string shellVerb)
      {
         // COMラッパーが初期化されていないか破棄されている場合のnullポインター逆参照を回避します。
         if (null == iQa || !iQa.IsValid)
         {
            return string.Empty;
         }

         // GetString()は例外をスローします。
         try
         {
            // この関数を2回呼び出すことを防ぐために大きなバッファを使用します。
            var size = NativeMethods.DefaultNativeQueryBufferSize;
            var buffer = new StringBuilder(size);

            iQa.GetString(Shell32.AssociationAttributes.NoTruncate | Shell32.AssociationAttributes.RemapRunDll, assocString, shellVerb, buffer, out size);

            return buffer.ToString();
         }
         catch
         {
            return string.Empty;
         }
      }


      [NonSerialized]
      private NativeMethods.QueryAssociationsWrapper _iQaNone;    // Shellから情報を取得。
      [NonSerialized]
      private NativeMethods.QueryAssociationsWrapper _iQaByExe;   // exeファイルから情報を取得。

      [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
      [SecurityCritical]
      private void Initialize()
      {
         if (Initialized)
         {
            return;
         }

         _iQaNone = NativeMethods.CreateQueryAssociations();

         if (_iQaNone.IsValid)
         {
            try
            {
               _iQaNone.Init(Shell32.AssociationAttributes.None, FullPath, 0, 0);

               _iQaByExe = NativeMethods.CreateQueryAssociations();

               if (_iQaByExe.IsValid)
               {
                  _iQaByExe.Init(Shell32.AssociationAttributes.InitByExeName, FullPath, 0, 0);

                  Initialized = true;
               }
            }
            catch
            {
            }
         }
      }


      /// <summary>オブジェクトの状態を更新します。</summary>
      [SecurityCritical]
      public void Refresh()
      {
         _iQaNone?.Dispose();
         _iQaByExe?.Dispose();
         _iQaNone = null;
         _iQaByExe = null;

         Association = Command = ContentType = DdeApplication = DefaultIcon = FriendlyAppName = FriendlyDocName = OpenWithAppName = null;
         Attributes = Shell32.GetAttributesOf.None;
         Initialized = false;
         Initialize();
      }


      /// <summary>Dispose 漏れ時のセーフネットとして基盤となる COM 参照を解放するファイナライザ。</summary>
      ~Shell32Info()
      {
         Dispose(false);
      }

      /// <summary>基盤となるCOM参照を解放します。</summary>
      public void Dispose()
      {
         Dispose(true);
         GC.SuppressFinalize(this);
      }

      private void Dispose(bool disposing)
      {
         _iQaNone?.Dispose();
         _iQaByExe?.Dispose();
         _iQaNone = null;
         _iQaByExe = null;
      }


      /// <summary>パスを文字列として返します。</summary>
      /// <returns>パス。</returns>
      public override string ToString()
      {
         return FullPath;
      }

      #endregion // Methods


      #region Properties

      private string _association;

      /// <summary>レジストリからShellのファイルまたはプロトコルの関連付けを取得します。</summary>
      public string Association
      {
         get
         {
            if (_association == null)
            {
               _association = GetString(_iQaNone, Shell32.AssociationString.Executable, null);
            }

            return _association;
         }

         private set { _association = value; }
      }

      
      private Shell32.GetAttributesOf _attributes;

      /// <summary>ファイルオブジェクトの属性。</summary>
      public Shell32.GetAttributesOf Attributes
      {
         get
         {
            if (_attributes == Shell32.GetAttributesOf.None)
            {
               var fileInfo = Shell32.GetFileInfoCore(FullPath, FileAttributes.Normal, Shell32.FileAttributes.Attributes, false, true);
               _attributes = fileInfo.Attributes;
            }

            return _attributes;
         }

         private set { _attributes = value; }
      }


      private string _command;

      /// <summary>レジストリからShellコマンドの関連付けを取得します。</summary>
      public string Command
      {
         get
         {
            if (_command == null)
            {
               _command = GetString(_iQaNone, Shell32.AssociationString.Command, null);
            }

            return _command;
         }

         private set { _command = value; }
      }


      private string _contentType;

      /// <summary>レジストリからShellコマンドの関連付けを取得します。</summary>
      public string ContentType
      {
         get
         {
            if (_contentType == null)
            {
               _contentType = GetString(_iQaNone, Shell32.AssociationString.ContentType, null);
            }

            return _contentType;
         }

         private set { _contentType = value; }
      }


      private string _ddeApplication;

      /// <summary>レジストリからShell DDEの関連付けを取得します。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dde")]
      public string DdeApplication
      {
         get
         {
            if (_ddeApplication == null)
            {
               _ddeApplication = GetString(_iQaNone, Shell32.AssociationString.DdeApplication, null);
            }

            return _ddeApplication;
         }

         private set { _ddeApplication = value; }
      }


      private string _defaultIcon;

      /// <summary>レジストリからShellのデフォルトアイコンの関連付けを取得します。</summary>
      public string DefaultIcon
      {
         get
         {
            if (_defaultIcon == null)
            {
               _defaultIcon = GetString(_iQaNone, Shell32.AssociationString.DefaultIcon, null);
            }

            return _defaultIcon;
         }

         private set { _defaultIcon = value; }
      }


      /// <summary>ファイルの完全修飾パスを表します。</summary>
      public string FullPath { get; private set; }


      private string _friendlyAppName;

      /// <summary>レジストリからShellのフレンドリーなアプリケーション名の関連付けを取得します。</summary>
      public string FriendlyAppName
      {
         get
         {
            if (_friendlyAppName == null)
            {
               _friendlyAppName = GetString(_iQaByExe, Shell32.AssociationString.FriendlyAppName, null);
            }

            return _friendlyAppName;
         }

         private set { _friendlyAppName = value; }
      }


      private string _friendlyDocName;

      /// <summary>レジストリからShellのフレンドリーなドキュメント名の関連付けを取得します。</summary>
      public string FriendlyDocName
      {
         get
         {
            if (_friendlyDocName == null)
            {
               _friendlyDocName = GetString(_iQaNone, Shell32.AssociationString.FriendlyDocName, null);
            }

            return _friendlyDocName;
         }

         private set { _friendlyDocName = value; }
      }


      /// <summary>インスタンスの初期化状態を反映します。</summary>
      internal bool Initialized { get; set; }


      private string _openWithAppName;

      /// <summary>レジストリからShellの「プログラムから開く」コマンドの関連付けを取得します。</summary>
      public string OpenWithAppName
      {
         get
         {
            if (_openWithAppName == null)
            {
               _openWithAppName = GetString(_iQaNone, Shell32.AssociationString.FriendlyAppName, null);
            }

            return _openWithAppName;
         }

         private set { _openWithAppName = value; }
      }

      #endregion // Properties
   }
}
