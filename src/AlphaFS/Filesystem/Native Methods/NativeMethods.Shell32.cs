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
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      #region AssocXxx

      /// <summary>IQueryAssociations オブジェクトへのポインタを返します。</summary>
      /// <returns>この関数が成功した場合、S_OK を返します。それ以外の場合、HRESULT エラーコードを返します。</returns>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional, Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint AssocCreate(Guid clsid, ref Guid riid, out IntPtr ppv);

      /// <summary>レジストリからファイルまたはプロトコルの関連付けに関連する文字列を検索して取得します。</summary>
      /// <returns>戻り値の型: HRESULT。S_OK、E_POINTER、S_FALSE を含む標準 COM エラー値を返します。</returns>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "AssocQueryStringW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint AssocQueryString(Shell32.AssociationAttributes flags, Shell32.AssociationString str, [MarshalAs(UnmanagedType.LPWStr)] string pszAssoc, [MarshalAs(UnmanagedType.LPWStr)] string pszExtra, StringBuilder pszOut, [MarshalAs(UnmanagedType.U4)] out uint pcchOut);


      #region IQueryAssociations

      internal static readonly Guid ClsidQueryAssociations = new Guid("A07034FD-6CAA-4954-AC3F-97A27216F98A");
      internal const string QueryAssociationsGuid = "C46CA590-3C3F-11D2-BEE6-0000F805CA57";

      /// <summary>IQueryAssociations COM インターフェイス (IUnknown ベース) の AOT セーフラッパー。
      /// ランタイム COM 相互運用の代わりに生の vtable 関数ポインタ呼び出しを使用します。</summary>
      internal sealed unsafe class QueryAssociationsWrapper : IDisposable
      {
         private nint _ptr;

         internal QueryAssociationsWrapper(nint comPtr)
         {
            _ptr = comPtr;
         }

         /// <summary>Dispose 漏れ時のセーフネットとして COM 参照を解放するファイナライザ。</summary>
         ~QueryAssociationsWrapper()
         {
            Dispose(false);
         }

         internal bool IsValid => _ptr != 0;

         /// <summary>IQueryAssociations インターフェイスを初期化し、ルートキーを適切な ProgID に設定します。</summary>
         internal void Init(Shell32.AssociationAttributes flags, string pszAssoc, nint hkProgid, nint hwnd)
         {
            // IUnknown vtable: [0] QueryInterface, [1] AddRef, [2] Release
            // IQueryAssociations vtable: [3] Init, [4] GetKey, [5] GetString, [6] GetData
            nint* vtable = *(nint**)_ptr;
            var initFn = (delegate* unmanaged[Stdcall]<nint, Shell32.AssociationAttributes, char*, nint, nint, int>)vtable[3];

            int hr;
            fixed (char* pAssoc = pszAssoc)
            {
               hr = initFn(_ptr, flags, pAssoc, hkProgid, hwnd);
            }

            Marshal.ThrowExceptionForHR(hr);
         }

         /// <summary>レジストリからファイルまたはプロトコルの関連付けに関連する文字列を検索して取得します。</summary>
         internal void GetString(Shell32.AssociationAttributes flags, Shell32.AssociationString str, string pwszExtra, StringBuilder pwszOut, out int pcchOut)
         {
            // IQueryAssociations vtable: [3] Init, [4] GetKey, [5] GetString
            nint* vtable = *(nint**)_ptr;
            var getStringFn = (delegate* unmanaged[Stdcall]<nint, Shell32.AssociationAttributes, Shell32.AssociationString, char*, char*, int*, int>)vtable[5];

            int hr;
            pcchOut = pwszOut.Capacity;
            // 出力用の一時バッファを割り当てる
            var buffer = new char[pcchOut];
            fixed (char* pExtra = pwszExtra)
            fixed (char* pOut = buffer)
            fixed (int* pSize = &pcchOut)
            {
               hr = getStringFn(_ptr, flags, str, pExtra, pOut, pSize);
            }

            if (hr >= 0)
            {
               pwszOut.Clear();
               pwszOut.Append(buffer, 0, pcchOut > 0 ? pcchOut - 1 : 0);
            }

            Marshal.ThrowExceptionForHR(hr);
         }

         /// <summary>COM オブジェクト参照を解放します。</summary>
         public void Dispose()
         {
            Dispose(true);
            GC.SuppressFinalize(this);
         }

         private void Dispose(bool disposing)
         {
            var ptr = _ptr;
            _ptr = 0;
            if (ptr != 0)
            {
               nint* vtable = *(nint**)ptr;
               var releaseFn = (delegate* unmanaged[Stdcall]<nint, uint>)vtable[2];
               releaseFn(ptr);
            }
         }
      }

      /// <summary>AssocCreate から新しい IQueryAssociations ラッパーを作成します。</summary>
      internal static QueryAssociationsWrapper CreateQueryAssociations()
      {
         var iid = new Guid(QueryAssociationsGuid);
         var hr = AssocCreate(ClsidQueryAssociations, ref iid, out var ptr);
         return hr == Win32Errors.S_OK ? new QueryAssociationsWrapper(ptr) : new QueryAssociationsWrapper(0);
      }

      #endregion // IQueryAssociations

      #endregion // AssocXxx


      #region Path

      /// <summary>ファイルやフォルダなどのファイルシステムオブジェクトへのパスが有効かどうかを判断します。</summary>
      /// <returns>ファイルが存在する場合は <c>true</c>、それ以外は <c>false</c>。拡張エラー情報を取得するには GetLastError を呼び出してください。</returns>
      /// <remarks>
      /// この関数はパスの有効性をテストします。
      /// UNC (汎用名前付け規則) で指定されたパスはファイルのみに制限されます。つまり、\\server\share\file は許可されます。
      /// サーバーまたはサーバー共有へのネットワーク共有パスは許可されません。つまり、\\server や \\server\share は不可です。
      /// マウントされたリモートドライブがサービス停止中の場合、この関数は FALSE を返します。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "PathFileExistsW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool PathFileExists([MarshalAs(UnmanagedType.LPWStr)] string pszPath);


      /// <summary>ファイル URL を Microsoft MS-DOS パスに変換します。</summary>
      /// <returns>型: HRESULT
      /// この関数が成功した場合、S_OK を返します。それ以外の場合、HRESULT エラーコードを返します。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional, Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "PathCreateFromUrlW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint PathCreateFromUrl([MarshalAs(UnmanagedType.LPWStr)] string pszUrl, StringBuilder pszPath, [MarshalAs(UnmanagedType.U4)] ref uint pcchPath, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>ファイル URL からパスを作成します。</summary>
      /// <returns>型: HRESULT
      /// この関数が成功した場合、S_OK を返します。それ以外の場合、HRESULT エラーコードを返します。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint PathCreateFromUrlAlloc([MarshalAs(UnmanagedType.LPWStr)] string pszIn, out StringBuilder pszPath, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>Microsoft MS-DOS パスを正規化された URL に変換します。</summary>
      /// <returns>型: HRESULT
      /// pszPath が既に URL 形式の場合、S_FALSE を返します。この場合、pszPath は単に pszUrl にコピーされます。
      /// それ以外の場合、成功すれば S_OK を返し、失敗した場合は標準 COM エラー値を返します。
      /// </returns>
      /// <remarks>
      /// UrlCreateFromPath は拡張パスをサポートしていません。これらは拡張長パスプレフィックス "\\?\" を含むパスです。
      /// </remarks>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional, Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "UrlCreateFromPathW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.U4)]
      internal static extern uint UrlCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, StringBuilder pszUrl, ref uint pcchUrl, [MarshalAs(UnmanagedType.U4)] uint dwFlags);


      /// <summary>URL が指定された種類かどうかをテストします。</summary>
      /// <returns>
      /// 型: BOOL
      /// URL の種類の1つを除くすべてについて、URL が指定された種類の場合 UrlIs は <c>true</c> を返し、それ以外は <c>true</c> を返します。
      /// UrlIs が <see cref="Shell32.UrlType.IsAppliable"/> に設定されている場合、UrlIs は URL スキームの判定を試みます。
      /// 関数がスキームを判定できた場合は <c>true</c> を返し、それ以外は <c>false</c> を返します。
      /// </returns>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional, Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "UrlIsW"), SuppressUnmanagedCodeSecurity]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool UrlIs([MarshalAs(UnmanagedType.LPWStr)] string pszUrl, Shell32.UrlType urlIs);

      #endregion // Path


      /// <summary>アイコンを破棄し、アイコンが占有していたメモリを解放します。</summary>
      /// <remarks>サポートされる最小クライアント: Windows XP [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("user32.dll", SetLastError = false)]
      [return: MarshalAs(UnmanagedType.Bool)]
      internal static extern bool DestroyIcon(IntPtr hIcon);


      /// <summary>ファイル、フォルダ、ディレクトリ、ドライブルートなど、ファイルシステム内のオブジェクトに関する情報を取得します。</summary>
      /// <remarks>この関数はバックグラウンドスレッドから呼び出す必要があります。そうしないと、UI が応答しなくなる可能性があります。</remarks>
      /// <remarks>サポートされる最小クライアント: Windows 2000 Professional [デスクトップアプリのみ]</remarks>
      /// <remarks>サポートされる最小サーバー: Windows 2000 Server [デスクトップアプリのみ]</remarks>
      [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage"), SuppressMessage("Microsoft.Security", "CA5122:PInvokesShouldNotBeSafeCriticalFxCopRule")]
      [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SHGetFileInfoW"), SuppressUnmanagedCodeSecurity]
      internal static extern IntPtr ShGetFileInfo([MarshalAs(UnmanagedType.LPWStr)] string pszPath, FileAttributes dwFileAttributes, [MarshalAs(UnmanagedType.Struct)] out Shell32.FileInfo psfi, [MarshalAs(UnmanagedType.U4)] uint cbFileInfo, [MarshalAs(UnmanagedType.U4)] Shell32.FileAttributes uFlags);
   }
}
